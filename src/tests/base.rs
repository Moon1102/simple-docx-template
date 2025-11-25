use crate::DOCX;
use crate::public::error::DocxError;
use crate::tests::base::PetType::{Cat, Dog};
use serde::Serialize;
use serde_json::{Value, json};
use std::collections::HashMap;
use tokio::fs::File as AsyncFile;
use tokio::io::{AsyncReadExt, BufReader};

#[derive(Serialize)]
struct User {
    name: String,
    age: u8,
    pets: Option<Vec<Pet>>,
}

#[derive(Serialize)]
struct Pet {
    name: String,
    toys: Option<Vec<Toy>>,
    r#type: PetType,
}

#[derive(Serialize)]
struct Toy {
    title: String,
    durability: f32,
    thumbnail: Option<String>,
}

#[derive(Serialize)]
enum PetType {
    Cat,
    Dog,
}

#[tokio::test]
async fn test_base() -> Result<(), DocxError> {
    let mut thumbnail = String::new();
    let mut logo = String::new();

    let mut reader = BufReader::new(
        AsyncFile::open("template/logo_base64.txt")
            .await
            .map_err(|e| DocxError::Xml(e.into()))?,
    );
    reader
        .read_to_string(&mut thumbnail)
        .await
        .map_err(|e| DocxError::Xml(e.into()))?;

    reader = BufReader::new(
        AsyncFile::open("template/thumbnail_base64.txt")
            .await
            .map_err(|e| DocxError::Xml(e.into()))?,
    );
    reader
        .read_to_string(&mut logo)
        .await
        .map_err(|e| DocxError::Xml(e.into()))?;

    let mut users = vec![];
    users.push(User {
        name: "Lisa".to_string(),
        age: 5,
        pets: None,
    });

    users.push(User {
        name: "Peter".to_string(),
        age: 7,
        pets: Some(vec![
            Pet {
                name: "Dog1".to_string(),
                toys: None,
                r#type: Dog,
            },
            Pet {
                name: "Cat1".to_string(),
                toys: Some(vec![
                    Toy {
                        title: "Doll1".to_string(),
                        durability: 59.99,
                        thumbnail: None,
                    },
                    Toy {
                        title: "Doll2".to_string(),
                        durability: 58.99,
                        thumbnail: None,
                    },
                ]),
                r#type: Cat,
            },
        ]),
    });

    users.push(User {
        name: "Adam".to_string(),
        age: 6,
        pets: Some(vec![Pet {
            name: "Dog2".to_string(),
            toys: Some(vec![Toy {
                title: "Doll3".to_string(),
                durability: 99.99,
                thumbnail: Some(thumbnail.clone()),
            }]),
            r#type: Dog,
        }]),
    });

    let users = users
        .iter()
        .filter_map(|u| serde_json::to_value(u).ok())
        .collect::<Vec<_>>();

    let mut data = HashMap::new();
    data.insert(
        "{{report title}}".to_string(),
        Value::String("New Title".to_string()),
    );
    data.insert(
        "{{report_subtitle}}".to_string(),
        Value::String("There is a sub_title".to_string()),
    );
    data.insert("{{report_logo}}".to_string(), Value::String(logo));

    let recorder = json!({"first_name": "Jim","last_name": "Green", "cnt": format!("{}",users.len()), "time": "2025-01-01 00:00:01","remark":"A very very very long stretch of meaningless text and a very very very long stretch of meaningless text and a very very very long stretch of meaningless text." });
    if let Some(map) = recorder.as_object() {
        for (k, v) in map {
            data.insert(format!("{{{{rec_{}}}}}", k), v.to_owned());
        }
    }

    let target =
        json!({"address":"85 The Vineyards","name": "Sam","photo":thumbnail,"city":"Chelmsford"});
    if let Some(map) = target.as_object() {
        for (k, v) in map {
            data.insert(format!("{{{{t_{}}}}}", k), v.to_owned());
        }
    }

    data.insert("{{#users}}".to_string(), Value::Array(users));

    let mut docx = DOCX::default();
    docx.generate("template/test.docx", "output/output.docx", &data)
        .await?;

    Ok(())
}
