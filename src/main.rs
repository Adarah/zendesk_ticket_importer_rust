use crate::importer::Importer;
use crate::objects::config::Config;
use anyhow::Result;
use std::path::PathBuf;
use structopt::StructOpt;
#[macro_use]
extern crate anyhow;
#[macro_use]
extern crate serde_json;

#[derive(StructOpt, Debug)]
#[structopt(name = "Zendesk Ticket Importer")]
pub struct Opt {
    /// Activate debug mode
    #[structopt(short, long)]
    debug: bool,

    /// Sets verbosity level (-v, -vv, -vvv)
    #[structopt(short, long, parse(from_occurrences))]
    verbose: u8,

    /// Input .xls, .xlsx, xlsb, or ods file
    #[structopt(name = "FILE", parse(from_os_str))]
    file: PathBuf,

    /// Toml file. Defaut location on linux is $HOME/.config/zendesk_ticket_importer
    #[structopt(short, long, name = "CONFIG_FILE")]
    config_path: Option<PathBuf>,
}

// TODO: Create installation script for linux and windows
#[tokio::main]
async fn main() -> Result<()> {
    let mut opt = Opt::from_args();
    let config = Config::from_opt(&mut opt)?;
    let importer = Importer::new(&opt.file, config)?;
    importer.run().await?;
    Ok(())
}

pub mod importer {
    use crate::objects::config::Config;
    use crate::objects::ticket::Ticket;
    use crate::objects::ticket::TicketWrapper;
    use anyhow::{Context, Result};
    use base64;
    use calamine::{self, open_workbook_auto, DataType, Range, Reader};
    use chrono::{DateTime, NaiveDate, Utc};
    use reqwest::{header, Client};
    use serde::{Deserialize, Serialize};
    use std::fmt::Debug;
    use std::path::Path;
    use std::time::Duration;

    pub struct Importer {
        range: Range<DataType>,
        config: Config,
        client: Client,
    }

    #[derive(Debug, Serialize, Deserialize)]
    pub struct GetFieldsReponse {
        ticket_fields: Vec<TicketField>,
    }

    #[derive(Debug, Serialize, Deserialize)]
    pub struct TicketField {
        pub id: usize,
        pub title: String,
        #[serde(rename = "type")]
        pub field_type: String,
        pub custom_field_options: Option<Vec<CustomField>>,
    }

    #[derive(Clone, Debug, Serialize, Deserialize)]
    pub struct CustomField {
        pub name: String,
        pub value: ApiValue,
    }

    #[derive(Clone, Debug, Serialize, Deserialize)]
    #[serde(untagged)]
    pub enum ApiValue {
        Common(String),
        DateTime(DateTime<Utc>),
        Date(NaiveDate),
        Checkbox(bool),
        MultiSelect(Vec<String>),
    }

    impl Importer {
        pub fn new<P>(file_path: P, config: Config) -> Result<Self>
        where
            P: AsRef<Path> + Debug + Copy,
        {
            let range = Importer::get_range(file_path, &config)?;
            let client = Importer::create_client(&config)?;
            return Ok(Importer {
                range,
                config,
                client,
            });
        }

        fn get_range<P>(file_path: P, config: &Config) -> Result<Range<DataType>>
        where
            P: AsRef<Path> + Debug + Copy,
        {
            let mut workbook = open_workbook_auto(file_path)
                .with_context(|| format!("Cannot open file: {:#?}", file_path))?;
            let s: Range<DataType> = workbook
                .worksheet_range(&config.worksheet.name)
                .with_context(|| format!("Could not find worksheet {:#?}", config.worksheet.name))?
                .unwrap(); // When does this fail? I will keep it as an unwrap for now
            return Ok(s);
        }

        fn create_client(config: &Config) -> Result<Client> {
            let authorization = base64::encode(format!(
                "{}/token:{}",
                config.credentials.email, config.credentials.api_token
            ));
            let mut headers = header::HeaderMap::new();
            headers.insert(
                header::AUTHORIZATION,
                header::HeaderValue::from_str(format!("Basic {}", authorization).as_str())?,
            );
            headers.insert(
                header::CONTENT_TYPE,
                header::HeaderValue::from_static("application/json"),
            );

            let client = Client::builder()
                .default_headers(headers)
                .timeout(Duration::new(10, 0))
                .use_rustls_tls()
                .build()?;
            return Ok(client);
        }

        pub async fn run(self) -> Result<()> {
            let api_fields = self.get_api_fields().await?;
            let mut tickets: Vec<Ticket> =
                Vec::with_capacity(self.range.rows().len() - self.config.worksheet.top_row - 1);
            for (row_num, row) in self
                .range
                .rows()
                .skip(self.config.worksheet.top_row - 1)
                .enumerate()
            {
                let ticket = Ticket::from_row(row, &self.config, &api_fields);
                if let Err(err) = ticket {
                    eprintln!(
                        "Error processing line {}, cause: {}",
                        row_num + self.config.worksheet.top_row,
                        err
                    );
                    continue;
                }
                tickets.push(ticket.unwrap());
            }
            // println!("{:#?}", tickets);
            for chunk in tickets.chunks(100) {
                let wrapper = TicketWrapper {
                    tickets: chunk.to_vec(),
                };
                // println!("{}", json!(&wrapper));
                let response = self
                    .client
                    .post(
                        format!(
                            "https://{}.zendesk.com{}",
                            self.config.credentials.subdomain, self.config.urls.post_many
                        )
                        .as_str(),
                    )
                    .json(&wrapper)
                    .send()
                    .await
                    .with_context(|| "Zendesk server didn't respond")?
                    .error_for_status()
                    .with_context(|| "The request for creating tickets failed")?;
                println!("{:#?}", response.text().await?);
            }
            Ok(())
        }

        pub async fn get_api_fields(&self) -> Result<Vec<TicketField>> {
            let fields_url = format!(
                "https://{}.zendesk.com{}",
                self.config.credentials.subdomain, self.config.urls.get_fields
            );
            let fields: GetFieldsReponse = self
                .client
                .get(&fields_url)
                .send()
                .await
                .with_context(|| "Zendesk server didn't respond")?
                .error_for_status()
                .with_context(|| "The request for custom url field ids returned an error")?
                .json()
                .await
                .with_context(|| "Could not parse response as json")?;
            // println!("{:#?}", fields);
            return Ok(fields.ticket_fields);
        }
    }
}

pub mod objects {
    pub mod config {

        use crate::objects::excel_mapper::TicketFields;
        use crate::Opt;
        use anyhow::{Context, Result};
        use serde::Deserialize;
        use std::path::PathBuf;
        use std::{env, fs};

        #[derive(Deserialize, Debug)]
        pub struct Config {
            #[serde(rename = "api_url")]
            pub urls: ApiUrls,
            pub credentials: Credentials,
            pub worksheet: Worksheet,
            pub ticket: TicketFields,
        }

        impl Config {
            pub fn from_opt(opt: &mut Opt) -> Result<Self> {
                if !opt.config_path.is_some() {
                    opt.config_path = Some(Config::get_default_path());
                }
                let content: String = fs::read_to_string(opt.config_path.as_ref().unwrap())
                    .with_context(|| "Could not read config file")?;
                return Ok(toml::from_str::<Self>(&content)
                    .with_context(|| "Failed to parse config file as toml")?);
            }

            fn get_default_path() -> PathBuf {
                match env::consts::OS {
                    "linux" => {
                        let home = env::var("HOME")
                            .expect("$HOME variable is not set!")
                            .clone();
                        return PathBuf::new()
                            .join(home)
                            .join(".config/zendesk_ticket_importer/config.toml");
                    }
                    "windows" => PathBuf::from("./config.toml"),
                    other_os => panic!("{} is not supported", other_os),
                }
            }
        }

        #[derive(Deserialize, Debug)]
        pub struct ApiUrls {
            pub get_fields: String,
            pub post_many: String,
        }

        #[derive(Deserialize, Debug)]
        pub struct Credentials {
            pub api_token: String,
            pub email: String,
            pub subdomain: String,
        }

        #[derive(Deserialize, Debug)]
        pub struct Worksheet {
            pub name: String,
            pub top_row: usize,
            pub timezone: String,
        }
    }

    mod excel_mapper {
        use serde::de::Error;
        use serde::{Deserialize, Deserializer};
        use std::collections::HashMap;

        fn excel_column_to_index(col_name: &str) -> Option<usize> {
            if col_name.is_empty() {
                return None;
            }
            let mut ans = 0;
            col_name.to_owned().make_ascii_uppercase();
            for (idx, letter) in col_name.chars().rev().enumerate() {
                let num = (letter as u32) - 64;
                ans += num * 26u32.pow(idx as u32);
            }
            return Some((ans - 1) as usize);
        }

        #[derive(Deserialize, Debug)]
        pub struct TicketFields {
            pub system_fields: SystemFields,
            #[serde(deserialize_with = "custom_flattener")]
            pub custom_fields: HashMap<String, usize>,
        }

        #[derive(Deserialize, Debug)]
        pub struct SystemFields {
            #[serde(deserialize_with = "custom_deserializer")]
            pub comment: usize,
            #[serde(deserialize_with = "custom_deserializer_opt")]
            pub subject: Option<usize>,
            #[serde(deserialize_with = "custom_deserializer_opt")]
            pub status: Option<usize>,
            #[serde(deserialize_with = "custom_deserializer_opt")]
            pub tickettype: Option<usize>,
            #[serde(deserialize_with = "custom_deserializer_opt")]
            pub assignee: Option<usize>,
            #[serde(deserialize_with = "custom_deserializer_opt")]
            pub priority: Option<usize>,
        }

        fn custom_flattener<'de, D>(deserializer: D) -> Result<HashMap<String, usize>, D::Error>
        where
            D: Deserializer<'de>,
        {
            let map: HashMap<&str, &str> = Deserialize::deserialize(deserializer)?;
            let mut new_map: HashMap<String, usize> = HashMap::with_capacity(map.capacity());
            for (k, v) in map.into_iter() {
                if v.is_empty() {
                    continue;
                }
                if !v.chars().all(|c| char::is_ascii_alphabetic(&c)) {
                    return Err(D::Error::custom(
                        "Excel columns can't contain non-ascii-aphabetic characters",
                    ));
                }
                new_map.insert(k.to_string(), excel_column_to_index(v).unwrap());
            }
            Ok(new_map)
        }
        fn custom_deserializer<'de, D>(deserializer: D) -> Result<usize, D::Error>
        where
            D: Deserializer<'de>,
        {
            let s: &str = Deserialize::deserialize(deserializer)?;

            if !s.chars().all(|c| char::is_ascii_alphabetic(&c)) {
                return Err(D::Error::custom(
                    "Excel columns can't contain non-ascii-aphabetic characters",
                ));
            }
            Ok(excel_column_to_index(s).unwrap())
        }

        fn custom_deserializer_opt<'de, D>(deserializer: D) -> Result<Option<usize>, D::Error>
        where
            D: Deserializer<'de>,
        {
            let s: &str = Deserialize::deserialize(deserializer)?;

            if !s.chars().all(|c| char::is_ascii_alphabetic(&c)) {
                return Err(D::Error::custom(
                    "Excel columns can't contain non-ascii-aphabetic characters",
                ));
            }
            Ok(excel_column_to_index(s))
        }
    }

    pub mod ticket {
        use crate::importer::{ApiValue, TicketField};
        use crate::objects::config::Config;
        use anyhow::Result;
        use calamine::DataType;
        use chrono::Utc;
        use chrono::{NaiveDateTime, TimeZone};
        use serde::Serialize;

        #[derive(Serialize, Debug)]
        pub struct TicketWrapper {
            pub tickets: Vec<Ticket>,
        }

        #[derive(Serialize, Debug, Clone)]
        pub struct Ticket {
            #[serde(skip_serializing_if = "Option::is_none")]
            subject: Option<String>,
            comment: Comment,
            #[serde(skip_serializing_if = "Option::is_none")]
            priority: Option<Priority>,
            #[serde(skip_serializing_if = "Option::is_none")]
            status: Option<Status>,
            #[serde(rename = "type", skip_serializing_if = "Option::is_none")]
            tickettype: Option<TicketType>,
            #[serde(skip_serializing_if = "Option::is_none")]
            assignee: Option<String>,
            custom_fields: Vec<Option<CustomFields>>,
        }

        impl Ticket {
            pub fn from_row(
                row: &[DataType],
                config: &Config,
                api_fields: &Vec<TicketField>,
            ) -> Result<Self> {
                let t = &config.ticket.system_fields;
                let subject = t
                    .subject
                    .and_then(|x| row.get(x))
                    .and_then(DataType::get_string)
                    .map(str::to_string);
                let comment = Comment {
                    body: row
                        .get(t.comment)
                        .and_then(DataType::get_string)
                        .ok_or_else(|| anyhow!("Comment cell should be a string"))?
                        .to_string(),
                };
                let priority = t
                    .priority
                    .and_then(|x| row.get(x))
                    .and_then(DataType::get_string)
                    .map(Priority::from_str)
                    .transpose()?;

                let status = t
                    .status
                    .and_then(|x| row.get(x))
                    .and_then(DataType::get_string)
                    .map(Status::from_str)
                    .transpose()?;
                let tickettype = t
                    .tickettype
                    .and_then(|x| row.get(x))
                    .and_then(DataType::get_string)
                    .map(TicketType::from_str)
                    .transpose()?;
                let assignee = t
                    .assignee
                    .and_then(|x| row.get(x))
                    .and_then(DataType::get_string)
                    .map(str::to_string);
                let mut custom_fields = Vec::with_capacity(config.ticket.custom_fields.len());
                for (key, value) in config.ticket.custom_fields.iter() {
                    let custom_field = api_fields
                        .iter()
                        .find(|&x| x.title == *key)
                        .map(|field| {
                            let data = row.get(*value);
                            match field.field_type.as_str() {
                                "integer" => CustomFields::from_integer(data, field),
                                "decimal" => CustomFields::from_decimal(data, field),
                                "date" => {
                                    CustomFields::from_date(data, field, &config.worksheet.timezone)
                                }
                                "checkbox" => CustomFields::from_checkbox(data, field),
                                "text" => CustomFields::from_text(data, field),
                                "tagger" => CustomFields::from_tagger(data, field),
                                unknown => {
                                    return Err(anyhow!("Unknown Zendesk field type: {}", unknown))
                                }
                            }
                        })
                        .transpose()?;
                    custom_fields.push(custom_field);
                }

                Ok(Ticket {
                    subject,
                    comment,
                    priority,
                    status,
                    tickettype,
                    assignee,
                    custom_fields,
                })
            }
        }

        #[derive(Serialize, Debug, Clone)]
        pub struct Comment {
            body: String,
        }

        #[derive(Serialize, Debug, Clone)]
        pub enum Priority {
            #[serde(rename = "low")]
            Low,
            #[serde(rename = "normal")]
            Normal,
            #[serde(rename = "high")]
            High,
            #[serde(rename = "urgent")]
            Urgent,
        }

        impl Priority {
            pub fn from_str(name: &str) -> Result<Self> {
                println!("convertendo priority");
                match name.to_ascii_lowercase().as_str() {
                    "low" | "baixa" => Ok(Priority::Low),
                    "normal" => Ok(Priority::Normal),
                    "high" | "alta" => Ok(Priority::High),
                    "urgent" | "urgente" => Ok(Priority::Urgent),
                    _ => Err(anyhow!(
                        "Unknown priority, expected low/baixa, normal, high/alta, urgente/urgente"
                    )),
                }
            }
        }

        #[derive(Serialize, Debug, Clone)]
        pub enum Status {
            #[serde(rename = "new")]
            New,
            #[serde(rename = "open")]
            Open,
            #[serde(rename = "pending")]
            Pending,
            #[serde(rename = "hold")]
            Hold,
            #[serde(rename = "solved")]
            Solved,
            #[serde(rename = "closed")]
            Closed,
        }

        impl Status {
            pub fn from_str(name: &str) -> Result<Self> {
                match name.to_ascii_lowercase().as_str() {
                    "open" | "aberto" => Ok(Status::Open),
                    "pending" | "pendente" => Ok(Status::Pending),
                    "hold" | "em espera" => Ok(Status::Hold),
                    "solved" | "resolvido" => Ok(Status::Solved),
                    "closed" | "fechado" => Ok(Status::Closed),
                    _ => Err(anyhow!("Unknown status, expected open/aberto, pending/pendente, hold/em espera, solved/resolvido, closed/fechado")),
                }
            }
        }

        #[derive(Serialize, Debug, Clone)]
        pub enum TicketType {
            #[serde(rename = "question")]
            Question,
            #[serde(rename = "incident")]
            Incident,
            #[serde(rename = "problem")]
            Problem,
            #[serde(rename = "task")]
            Task,
        }

        impl TicketType {
            pub fn from_str(name: &str) -> Result<Self> {
                match name.to_ascii_lowercase().as_str() {
                    "question" | "pergunta" => Ok(TicketType::Question),
                    "incident" | "incidente" => Ok(TicketType::Incident),
                    "problem" | "problema" => Ok(TicketType::Problem),
                    "task" | "tarefa" => Ok(TicketType::Task),
                    _ => Err(anyhow!("Unknown ticket type, expected question/pergunta, incident/incidente, problem/problema, task/tarefa")),
                }
            }
        }

        #[derive(Serialize, Debug, Clone)]
        pub struct CustomFields {
            id: usize,
            value: ApiValue,
        }

        impl CustomFields {
            pub fn from_integer(
                excel_data: Option<&DataType>,
                api_field: &TicketField,
            ) -> Result<Self> {
                let id = api_field.id;
                let value = excel_data
                    // .and_then(DataType::get_int)
                    .and_then(DataType::get_float) // apparently even whole numbers are stored as floats? we should look into this later
                    .map(|x| x.to_string())
                    .map(ApiValue::Common)
                    .ok_or(anyhow!("Could not parse {:#?} as integer", excel_data));
                Ok(Self { id, value: value? })
            }

            pub fn from_decimal(
                excel_data: Option<&DataType>,
                api_field: &TicketField,
            ) -> Result<Self> {
                let id = api_field.id;
                let value = excel_data
                    .and_then(DataType::get_float)
                    .map(|x| x.to_string())
                    .map(ApiValue::Common)
                    .ok_or(anyhow!("Could not parse cell as float"));
                Ok(Self { id, value: value? })
            }

            pub fn from_checkbox(
                excel_data: Option<&DataType>,
                api_field: &TicketField,
            ) -> Result<Self> {
                let id = api_field.id;
                let value = excel_data
                    .and_then(DataType::get_bool)
                    .map(ApiValue::Checkbox)
                    .ok_or(anyhow!("Could not parse cell as boolean"));
                Ok(Self { id, value: value? })
            }

            pub fn from_text(
                excel_data: Option<&DataType>,
                api_field: &TicketField,
            ) -> Result<Self> {
                let id = api_field.id;
                let value = excel_data
                    .and_then(DataType::get_string)
                    .map(str::to_string)
                    .map(ApiValue::Common)
                    .ok_or(anyhow!("Could not parse cell as string"));
                Ok(Self { id, value: value? })
            }

            pub fn from_date(
                excel_data: Option<&DataType>,
                api_field: &TicketField,
                timezone: &str,
            ) -> Result<Self> {
                let zone = match timezone {
                    "Acre" => chrono_tz::Brazil::Acre,
                    "DeNoronha" => chrono_tz::Brazil::DeNoronha,
                    "East" => chrono_tz::Brazil::East,
                    "West" => chrono_tz::Brazil::West,
                    _ => return Err(anyhow!("Unknown timezone")),
                };

                let id = api_field.id;
                let value = excel_data
                    .and_then(DataType::get_float)
                    .map(|x| (x - 25569_f64) * 86400_f64)
                    .map(|x| NaiveDateTime::from_timestamp(x.round() as i64, 0))
                    .map(|x| zone.from_local_datetime(&x).unwrap())
                    .map(|x| x.with_timezone(&Utc))
                    .map(|x| x.date().naive_utc())
                    .map(ApiValue::Date)
                    .ok_or(anyhow!("Could not parse {:#?} as datetime", excel_data));
                Ok(Self { id, value: value? })
            }

            pub fn from_tagger(
                excel_data: Option<&DataType>,
                api_field: &TicketField,
            ) -> Result<Self> {
                let id = api_field.id;
                let value = excel_data
                    .and_then(DataType::get_string)
                    .and_then(|xcl| {
                        api_field
                            .custom_field_options
                            .as_ref()
                            .and_then(|v| v.into_iter().find(|field| field.name == xcl))
                    })
                    .map(|x| x.value.clone())
                    .ok_or(anyhow!("Could not parse cell as dropdown menu option"));
                Ok(Self { id, value: value? })
            }
        }
    }
}
