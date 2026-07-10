# CRUD Test Cases

機能拡張時のラバーダック用テストケースです。
変換定義とSQLを固定し、解析結果と将来の `アウトプット` シート出力を比較しながら実装します。

## 変換定義

| 所属テーブルID | 所属テーブル和名 | フィールドID | フィールド和名 |
| --- | --- | --- | --- |
| users | ユーザー | user_id | ユーザーID |
| users | ユーザー | name | 氏名 |
| users | ユーザー | email | メール |
| orders | 注文 | order_id | 注文ID |
| orders | 注文 | user_id | 注文ユーザーID |
| orders | 注文 | amount | 金額 |
| order_items | 注文明細 | order_id | 明細注文ID |
| order_items | 注文明細 | product_id | 商品ID |
| order_items | 注文明細 | quantity | 数量 |
| - |  | status | 状態 |
| - |  | created_at | 作成日時 |
| - |  | updated_at | 更新日時 |

## SELECT

入力:

```sql
select trim(users.name) as name, users.user_id, orders.order_id, status
from users
inner join orders on users.user_id = orders.user_id
where status = 'ACTIVE'
```

期待する和名変換後:

```sql
select trim(users.氏名) as name, users.ユーザーID, orders.注文ID, 状態
from users
inner join orders on users.ユーザーID = orders.注文ユーザーID
where 状態 = 'ACTIVE'
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.氏名 |
| 2 | users.ユーザーID |
| 3 | orders.注文ID |
| 4 | 状態 |
| 5 | orders.注文ユーザーID |

## INSERT

入力:

```sql
insert into orders (order_id, user_id, amount, status, created_at)
select orders.order_id, users.user_id, orders.amount, status, created_at
from users
```

期待する和名変換後:

```sql
insert into orders (order_id, user_id, amount, 状態, 作成日時)
select orders.注文ID, users.ユーザーID, orders.金額, 状態, 作成日時
from users
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | 状態 |
| 2 | 作成日時 |
| 3 | orders.注文ID |
| 4 | users.ユーザーID |
| 5 | orders.金額 |

## UPDATE

入力:

```sql
update users
set users.name = 'Taro',
    updated_at = CURRENT_TIMESTAMP,
    status = 'ACTIVE'
where users.user_id = :user_id
```

期待する和名変換後:

```sql
update users
set users.氏名 = 'Taro',
    更新日時 = CURRENT_TIMESTAMP,
    状態 = 'ACTIVE'
where users.ユーザーID = :user_id
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.氏名 |
| 2 | 更新日時 |
| 3 | 状態 |
| 4 | users.ユーザーID |

## DELETE

入力:

```sql
delete from orders
where orders.order_id in (
    select order_items.order_id
    from order_items
    where order_items.product_id = :product_id
)
and status = 'CANCELLED'
```

期待する和名変換後:

```sql
delete from orders
where orders.注文ID in (
    select order_items.明細注文ID
    from order_items
    where order_items.商品ID = :product_id
)
and 状態 = 'CANCELLED'
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ID |
| 2 | order_items.明細注文ID |
| 3 | order_items.商品ID |
| 4 | 状態 |

## 出力シート

将来の整形出力先シート名は `アウトプット` です。
現時点ではシートの存在だけをテストし、具体的な列構成は次の仕様検討で決めます。
