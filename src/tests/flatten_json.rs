fn get_test_data() -> serde_json::Value {
    serde_json::json!({
        "name": "Peter",
        "age": 7,
        "pets": [
            {
                "name": "Dog1",
                "toy": serde_json::Value::Null,
                "type": "Dog"
            },
            {
                "name": "Cat1",
                "toy": {
                    "title": "Doll1",
                    "durability": 59.99,
                    "thumbnail": serde_json::Value::Null
                },
                "type": "Cat"
            }
        ]
    })
}

#[test]
fn test_flatten_json_0() {
    use serde_json::json;

    let result = crate::core::utils::flatten_json(&get_test_data());

    assert_eq!(result.len(), 2);
    assert_eq!(result[0].get("name"), Some(&json!("Peter")));
    assert_eq!(result[0].get("pets.name"), Some(&json!("Dog1")));
    assert_eq!(result[1].get("pets.toy.title"), Some(&json!("Doll1")));
}

#[test]
fn test_flatten_json_1() {
    use serde_json::json;

    let result = crate::core::utils::flatten_json(&get_test_data());

    println!("Items：");
    for (i, record) in result.iter().enumerate() {
        println!("item {}: {:?}", i + 1, record);
    }

    // 验证结果
    assert_eq!(result.len(), 2);

    // 第一个记录 - Dog
    assert_eq!(result[0].get("name"), Some(&json!("Peter")));
    assert_eq!(result[0].get("age"), Some(&json!(7)));
    assert_eq!(result[0].get("pets.name"), Some(&json!("Dog1")));
    assert_eq!(result[0].get("pets.type"), Some(&json!("Dog")));
    assert_eq!(result[0].get("pets.toy"), Some(&serde_json::Value::Null));

    // 第二个记录 - Cat
    assert_eq!(result[1].get("name"), Some(&json!("Peter")));
    assert_eq!(result[1].get("age"), Some(&json!(7)));
    assert_eq!(result[1].get("pets.name"), Some(&json!("Cat1")));
    assert_eq!(result[1].get("pets.type"), Some(&json!("Cat")));
    assert_eq!(result[1].get("pets.toy.title"), Some(&json!("Doll1")));
    assert_eq!(result[1].get("pets.toy.durability"), Some(&json!(59.99)));
    assert_eq!(
        result[1].get("pets.toy.thumbnail"),
        Some(&serde_json::Value::Null)
    );
}
