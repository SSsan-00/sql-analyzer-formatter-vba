# CRUD Test Cases

機能拡張時のラバーダック用テストケースです。
変換定義とSQLを固定し、解析結果と将来の `アウトプット` シート出力を比較しながら実装します。

解析対象SQLは、A5:SQL Mk-2（A5M2）の `Ctrl+q` で整形した結果を `SQL解析` シートのA列セルへ貼り付ける前提です。
テストコードではA5M2が付ける行末スペースを `TS(...)` で明示して保持します。この文書では視認性のため行末スペースを省略しています。

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
| users | ユーザー | manager_id | 管理者ID |
| manager | 管理者 | user_id | ユーザーID |
| manager | 管理者 | name | 氏名 |
| manager | 管理者 | status | 状態 |
| - |  | status | 状態 |
| - |  | created_at | 作成日時 |
| - |  | updated_at | 更新日時 |

## SELECT

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    trim(users.name) as name
    , users.user_id
    , orders.order_id
    , status
from
    users
    inner join orders
        on users.user_id = orders.user_id
where
    status = 'ACTIVE'
```

期待する和名変換後

```sql
select
    trim(users.氏名) as name
    , users.ユーザーID
    , orders.注文ID
    , 状態
from
    users
    inner join orders
        on users.ユーザーID = orders.注文ユーザーID
where
    状態 = 'ACTIVE'
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

入力（A5M2 Ctrl+Q 整形後）

```sql
insert
into orders(order_id, user_id, amount, status, created_at)
select
    orders.order_id
    , users.user_id
    , orders.amount
    , status
    , created_at
from
    users
```

期待する和名変換後

```sql
insert
into orders(order_id, user_id, amount, 状態, 作成日時)
select
    orders.注文ID
    , users.ユーザーID
    , orders.金額
    , 状態
    , 作成日時
from
    users
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

入力（A5M2 Ctrl+Q 整形後）

```sql
update users
set
    users.name = 'Taro'
    , updated_at = CURRENT_TIMESTAMP
    , status = 'ACTIVE'
where
    users.user_id = :user_id
```

期待する和名変換後

```sql
update users
set
    users.氏名 = 'Taro'
    , 更新日時 = CURRENT_TIMESTAMP
    , 状態 = 'ACTIVE'
where
    users.ユーザーID = :user_id
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.氏名 |
| 2 | 更新日時 |
| 3 | 状態 |
| 4 | users.ユーザーID |

## DELETE

入力（A5M2 Ctrl+Q 整形後）

```sql
delete
from
    orders
where
    orders.order_id in (
        select
            order_items.order_id
        from
            order_items
        where
            order_items.product_id = :product_id
    )
    and status = 'CANCELLED'
```

期待する和名変換後

```sql
delete
from
    orders
where
    orders.注文ID in (
        select
            order_items.明細注文ID
        from
            order_items
        where
            order_items.商品ID = :product_id
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

## 複合SELECT CASE / WHEN / ELSE / HAVING / ORDER BY

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , case
        when sum(orders.amount) > 100000
            then 'VIP'
        when sum(orders.amount) between 50000 and 100000
            then 'STANDARD'
        else status
        end as rank_name
from
    users
    left join orders
        on users.user_id = orders.user_id
where
    (
        (status = 'ACTIVE' and orders.amount > 0)
        or (
            status = 'PENDING'
            and exists (
                select
                    1
                from
                    order_items
                where
                    order_items.order_id = orders.order_id
                    and order_items.quantity > 1
            )
        )
    )
group by
    users.user_id
    , status
having
    count(orders.order_id) > 0
order by
    users.user_id
    , status
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , case
        when sum(orders.金額) > 100000
            then 'VIP'
        when sum(orders.金額) between 50000 and 100000
            then 'STANDARD'
        else 状態
        end as rank_name
from
    users
    left join orders
        on users.ユーザーID = orders.注文ユーザーID
where
    (
        (状態 = 'ACTIVE' and orders.金額 > 0)
        or (
            状態 = 'PENDING'
            and exists (
                select
                    1
                from
                    order_items
                where
                    order_items.明細注文ID = orders.注文ID
                    and order_items.数量 > 1
            )
        )
    )
group by
    users.ユーザーID
    , 状態
having
    count(orders.注文ID) > 0
order by
    users.ユーザーID
    , 状態
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | orders.金額 |
| 3 | 状態 |
| 4 | orders.注文ユーザーID |
| 5 | order_items.明細注文ID |
| 6 | orders.注文ID |
| 7 | order_items.数量 |

## 自己結合

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , users.name
    , manager.name as manager_name
from
    users
    inner join users manager
        on users.manager_id = manager.user_id
where
    manager.status = status
order by
    manager.name
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , users.氏名
    , manager.氏名 as manager_name
from
    users
    inner join users manager
        on users.管理者ID = manager.ユーザーID
where
    manager.状態 = 状態
order by
    manager.氏名
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | users.氏名 |
| 3 | manager.氏名 |
| 4 | users.管理者ID |
| 5 | manager.ユーザーID |
| 6 | manager.状態 |
| 7 | 状態 |

## SELECT-INTO

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , users.email
    , status
into user_export
from
    users
where
    users.email is not null
    and status in ('ACTIVE', 'LOCKED')
order by
    users.email
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , users.メール
    , 状態
into user_export
from
    users
where
    users.メール is not null
    and 状態 in ('ACTIVE', 'LOCKED')
order by
    users.メール
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | users.メール |
| 3 | 状態 |

## UPDATE-FROM

入力（A5M2 Ctrl+Q 整形後）

```sql
update orders
set
    orders.amount = orders.amount * 1.1
    , updated_at = CURRENT_TIMESTAMP
from
    orders
    inner join users
        on orders.user_id = users.user_id
where
    (users.email like :domain or status = 'PENDING')
    and (
        orders.amount > 1000
        or exists (
            select
                1
            from
                order_items
            where
                order_items.order_id = orders.order_id
        )
    )
```

期待する和名変換後

```sql
update orders
set
    orders.金額 = orders.金額 * 1.1
    , 更新日時 = CURRENT_TIMESTAMP
from
    orders
    inner join users
        on orders.注文ユーザーID = users.ユーザーID
where
    (users.メール like :domain or 状態 = 'PENDING')
    and (
        orders.金額 > 1000
        or exists (
            select
                1
            from
                order_items
            where
                order_items.明細注文ID = orders.注文ID
        )
    )
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.金額 |
| 2 | 更新日時 |
| 3 | orders.注文ユーザーID |
| 4 | users.ユーザーID |
| 5 | users.メール |
| 6 | 状態 |
| 7 | order_items.明細注文ID |
| 8 | orders.注文ID |

## DELETE EXISTS

入力（A5M2 Ctrl+Q 整形後）

```sql
delete
from
    order_items
where
    exists (
        select
            1
        from
            orders
        where
            orders.order_id = order_items.order_id
            and (status = 'CANCELLED' or orders.amount <= 0)
    )
```

期待する和名変換後

```sql
delete
from
    order_items
where
    exists (
        select
            1
        from
            orders
        where
            orders.注文ID = order_items.明細注文ID
            and (状態 = 'CANCELLED' or orders.金額 <= 0)
    )
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ID |
| 2 | order_items.明細注文ID |
| 3 | 状態 |
| 4 | orders.金額 |

## T-SQL独立ケース

T-SQLで使用する関数・構文ごとに、A5M2 `Ctrl+q` 整形後の入力と和名変換後を独立した行として検証します。
`FORMAT` は複数のフォーマット指定子、`CAST` は複数の型を同一ケース内にまとめています。

### TRIM FROM

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    trim('.' from users.name) as trimmed_name
from
    users
where
    users.user_id = @user_id
```

期待する和名変換後

```sql
select
    trim('.' from users.氏名) as trimmed_name
from
    users
where
    users.ユーザーID = @user_id
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.氏名 |
| 2 | users.ユーザーID |

### IN

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , status
from
    users
where
    status in ('ACTIVE', 'LOCKED', 'PENDING')
    and users.user_id in (
        select
            orders.user_id
        from
            orders
        where
            orders.amount > 0
    )
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , 状態
from
    users
where
    状態 in ('ACTIVE', 'LOCKED', 'PENDING')
    and users.ユーザーID in (
        select
            orders.注文ユーザーID
        from
            orders
        where
            orders.金額 > 0
    )
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | 状態 |
| 3 | orders.注文ユーザーID |
| 4 | orders.金額 |

### COALESCE

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , coalesce(users.email, users.name, 'unknown') as contact_text
from
    users
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , coalesce(users.メール, users.氏名, 'unknown') as contact_text
from
    users
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | users.メール |
| 3 | users.氏名 |

### FORMAT

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    orders.order_id
    , format(orders.amount, 'N2', 'ja-JP') as amount_n2
    , format(created_at, 'yyyy/MM/dd') as created_date
    , format(created_at, 'yyyyMMddHHmmss') as created_stamp
    , format(orders.amount, 'C', 'ja-JP') as amount_currency
from
    orders
```

期待する和名変換後

```sql
select
    orders.注文ID
    , format(orders.金額, 'N2', 'ja-JP') as amount_n2
    , format(作成日時, 'yyyy/MM/dd') as created_date
    , format(作成日時, 'yyyyMMddHHmmss') as created_stamp
    , format(orders.金額, 'C', 'ja-JP') as amount_currency
from
    orders
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ID |
| 2 | orders.金額 |
| 3 | 作成日時 |

### WITH

入力（A5M2 Ctrl+Q 整形後）

```sql
with target_users as (
    select
        users.user_id
    from
        users
    where
        status = 'ACTIVE'
)
select
    target_users.user_id
from
    target_users
```

期待する和名変換後

```sql
with target_users as (
    select
        users.ユーザーID
    from
        users
    where
        状態 = 'ACTIVE'
)
select
    target_users.user_id
from
    target_users
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | 状態 |

### CAST

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    cast(users.user_id as int) as user_id_int
    , cast(orders.amount as decimal (18, 2)) as amount_decimal
    , cast(created_at as date) as created_date
    , cast(updated_at as datetime2(3)) as updated_at_dt
    , cast(status as nvarchar(20)) as status_text
from
    users
    inner join orders
        on users.user_id = orders.user_id
```

期待する和名変換後

```sql
select
    cast(users.ユーザーID as int) as user_id_int
    , cast(orders.金額 as decimal (18, 2)) as amount_decimal
    , cast(作成日時 as date) as created_date
    , cast(更新日時 as datetime2(3)) as updated_at_dt
    , cast(状態 as nvarchar(20)) as status_text
from
    users
    inner join orders
        on users.ユーザーID = orders.注文ユーザーID
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | orders.金額 |
| 3 | 作成日時 |
| 4 | 更新日時 |
| 5 | 状態 |
| 6 | orders.注文ユーザーID |

### ISNULL

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , isnull(users.email, 'unknown') as email_text
    , isnull(status, 'UNKNOWN') as status_text
from
    users
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , isnull(users.メール, 'unknown') as email_text
    , isnull(状態, 'UNKNOWN') as status_text
from
    users
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | users.メール |
| 3 | 状態 |

### SUBSTRING

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , substring(users.email, 1, 3) as email_prefix
from
    users
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , substring(users.メール, 1, 3) as email_prefix
from
    users
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | users.メール |

### ROUND

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    orders.order_id
    , round(orders.amount, 0) as amount_round0
    , round(orders.amount, 2, 1) as amount_truncate2
from
    orders
```

期待する和名変換後

```sql
select
    orders.注文ID
    , round(orders.金額, 0) as amount_round0
    , round(orders.金額, 2, 1) as amount_truncate2
from
    orders
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ID |
| 2 | orders.金額 |

### SUM

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    orders.user_id
    , sum(orders.amount) as total_amount
from
    orders
group by
    orders.user_id
having
    sum(orders.amount) > 0
```

期待する和名変換後

```sql
select
    orders.注文ユーザーID
    , sum(orders.金額) as total_amount
from
    orders
group by
    orders.注文ユーザーID
having
    sum(orders.金額) > 0
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ユーザーID |
| 2 | orders.金額 |

### REPLACE

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , replace (users.email, '@old.example', '@new.example') as normalized_email
from
    users
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , replace (users.メール, '@old.example', '@new.example') as normalized_email
from
    users
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | users.メール |

### DATEADD

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    orders.order_id
    , dateadd(day, 7, created_at) as due_date
    , dateadd(month, 1, created_at) as next_month_date
from
    orders
```

期待する和名変換後

```sql
select
    orders.注文ID
    , dateadd(day, 7, 作成日時) as due_date
    , dateadd(month, 1, 作成日時) as next_month_date
from
    orders
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ID |
| 2 | 作成日時 |

### DATEDIFF

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    orders.order_id
    , datediff(day, created_at, updated_at) as elapsed_days
    , datediff(minute, created_at, updated_at) as elapsed_minutes
from
    orders
```

期待する和名変換後

```sql
select
    orders.注文ID
    , datediff(day, 作成日時, 更新日時) as elapsed_days
    , datediff(minute, 作成日時, 更新日時) as elapsed_minutes
from
    orders
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ID |
| 2 | 作成日時 |
| 3 | 更新日時 |

### COUNT

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , count(orders.order_id) as order_count
from
    users
    left join orders
        on users.user_id = orders.user_id
group by
    users.user_id
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , count(orders.注文ID) as order_count
from
    users
    left join orders
        on users.ユーザーID = orders.注文ユーザーID
group by
    users.ユーザーID
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | orders.注文ID |
| 3 | orders.注文ユーザーID |

### EXISTS

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
from
    users
where
    exists (
        select
            1
        from
            orders
        where
            orders.user_id = users.user_id
            and orders.amount > 0
    )
```

期待する和名変換後

```sql
select
    users.ユーザーID
from
    users
where
    exists (
        select
            1
        from
            orders
        where
            orders.注文ユーザーID = users.ユーザーID
            and orders.金額 > 0
    )
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | orders.注文ユーザーID |
| 3 | orders.金額 |

## 出力シート

整形出力先シート名は `アウトプット` です。
現時点では暫定形式として、A列へ和名変換後クエリをもとにしたクエリブロックを出力します。
WITH句やサブクエリを含む場合は、内側のサブクエリ、外側のサブクエリ、解析対象のクエリ全体の順で出力します。
未対応のクエリは分解せず、そのまま1ブロックとして出力します。
