# วิธีใช้งาน Free Functions

โดย [ThepExcel.com](https://www.thepexcel.com/thepexcel-mfx/)

## วิธีติดตั้ง

1. เปิด Power Query Editor ใน Excel หรือ Power BI
2. สร้าง Blank Query ใหม่
3. เปิด Advanced Editor แล้ววางโค้ดจากไฟล์ `.pq` ที่ต้องการ
4. ตั้งชื่อ Query ให้ตรงกับชื่อฟังก์ชัน (เช่น `ThepTrim`)
5. เรียกใช้ฟังก์ชันจาก Query อื่นได้เลยครับ

---

## สารบัญ

### Text Processing (จัดการข้อความ)
- [ThepTrim](#theptrim) — ตัดช่องว่างซ้ำ
- [ThepExtractNumber](#thepextractnumber) — ดึงเฉพาะตัวเลขจากข้อความ
- [ThepThaiNumbertoArabic](#thepthainumbertoarabic) — แปลงเลขไทยเป็นเลขอารบิก

### Regex (Regular Expression)
- [ThepRegExExtract](#thepregexextract) — ดึงข้อความที่ match regex (ตัวแรก)
- [ThepRegExReplace](#thepregexreplace) — แทนที่ด้วย regex
- [ThepRegExMatchCount](#thepregexmatchcount) — นับจำนวน match

### Date (จัดการวันที่)
- [ThepDatefromText](#thepdatefromtext) — แปลง text เป็น date ตาม format ที่กำหนด
- [ThepGenDateTableFromDate](#thepgendatetablefromdate) — สร้างตารางวันที่จาก date
- [ThepGenDateTableFromText](#thepgendatetablefromtext) — สร้างตารางวันที่จาก text
- [ThepNETWORKDAYS](#thepnetworkdays) — นับวันทำการ

### Table Operations (จัดการตาราง)
- [ThepGetColumnName](#thepgetcolumnname) — ดึงชื่อคอลัมน์จากลำดับ
- [ThepGetMultipleListItem](#thepgetmultiplelistitem) — ดึงหลาย item จาก list
- [ThepRenameColumn](#theprenamecolumn) — เปลี่ยนชื่อคอลัมน์ด้วยเลขลำดับ
- [ThepReplaceAllError](#thepreplaceall error) — แทนที่ error ทุกคอลัมน์
- [ThepOneHot](#theponehot) — One-Hot Encoding

### API
- [ThepOpenAI](#thepopenai) — เรียก OpenAI API (เวอร์ชัน Free)

---

## Text Processing (จัดการข้อความ)

### ThepTrim

ตัดช่องว่างซ้ำ (หรือตัวอักษรอื่นที่ซ้ำ) ให้เหลือตัวเดียวครับ ทำงานเหมือน `TRIM` ของ Excel แต่ระบุตัวอักษรที่จะตัดได้

```
ThepTrim(OriginalText as text, optional TrimChar as text) as text
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalText` | text | ข้อความต้นฉบับ |
| `TrimChar` | text (optional) | ตัวอักษรที่ต้องการตัด (default: `" "` ช่องว่าง) |

#### ตัวอย่าง

```
ThepTrim("Hello     World")
// ผลลัพธ์: "Hello World"

ThepTrim("A---B---C", "-")
// ผลลัพธ์: "A-B-C"

// ใช้ใน Transform Column
= Table.TransformColumns(MyTable, {{"Name", each ThepTrim(_)}})
```

---

### ThepExtractNumber

ดึงเฉพาะตัวเลขจากข้อความ เลือกเก็บจุดทศนิยมหรือช่องว่างได้ครับ

```
ThepExtractNumber(OriginalText as text, optional Keepdot as logical, optional Keepspace as logical) as text
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalText` | text | ข้อความต้นฉบับ |
| `Keepdot` | logical (optional) | เก็บจุด `.` ไว้ไหม (default: `true`) |
| `Keepspace` | logical (optional) | เก็บช่องว่างไว้ไหม (default: `true`) |

#### ตัวอย่าง

```
ThepExtractNumber("ราคา 1,234.50 บาท")
// ผลลัพธ์: "1234.50"

ThepExtractNumber("โทร 081-234-5678")
// ผลลัพธ์: "081 234 5678" (เก็บ space)

ThepExtractNumber("ID: ABC-123", false, false)
// ผลลัพธ์: "123" (ไม่เก็บจุดหรือ space)

// ถ้าต้องการเป็น number ให้ครอบด้วย Number.From
= Number.From(ThepExtractNumber([Price]))
```

---

### ThepThaiNumbertoArabic

แปลงเลขไทย (๐-๙) เป็นเลขอารบิก (0-9) ครับ ตัวอักษรอื่นที่ไม่ใช่เลขไทยจะคงเดิม

```
ThepThaiNumbertoArabic(OriginalText as text) as text
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalText` | text | ข้อความที่มีเลขไทย |

#### ตัวอย่าง

```
ThepThaiNumbertoArabic("แมว ๓๘ ตัว ราคา ๕๔๓๖ บาท")
// ผลลัพธ์: "แมว 38 ตัว ราคา 5436 บาท"

ThepThaiNumbertoArabic("๐๘๑-๒๓๔-๕๖๗๘")
// ผลลัพธ์: "081-234-5678"
```

---

## Regex (Regular Expression)

> ฟังก์ชัน Regex ทุกตัวใช้ JavaScript ผ่าน `Web.Page()` ทำงานได้เฉพาะ **Excel Desktop** เท่านั้น ไม่รองรับ Power BI Service ครับ

### ThepRegExExtract

ดึงข้อความที่ตรงกับ Regex pattern (ตัวแรก หรือทั้งหมดถ้าใส่ `"g"`) คืนเป็น text หรือ list

```
ThepRegExExtract(OriginalText as text, RegExPattern as text, optional RegExMode as text) as any
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalText` | text | ข้อความต้นฉบับ |
| `RegExPattern` | text | Regex pattern |
| `RegExMode` | text (optional) | `"g"` = ทั้งหมด, `"i"` = ไม่สนตัวพิมพ์, `"gi"` = ทั้งสอง (default: ตัวแรก, case sensitive) |

#### ตัวอย่าง

```
// ดึง match แรก
ThepRegExExtract("abc 123 def 456", "\d+")
// ผลลัพธ์: "123"

// ดึงทั้งหมด
ThepRegExExtract("abc 123 def 456", "\d+", "g")
// ผลลัพธ์: {"123", "456"}

// Case insensitive
ThepRegExExtract("Hello World", "hello", "i")
// ผลลัพธ์: "Hello"
```

---

### ThepRegExReplace

แทนที่ข้อความด้วย Regex pattern

```
ThepRegExReplace(OriginalText as text, RegExPattern as text, NewText as text, optional RegExMode as text) as any
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalText` | text | ข้อความต้นฉบับ |
| `RegExPattern` | text | Regex pattern ที่ต้องการแทนที่ |
| `NewText` | text | ข้อความใหม่ (รองรับ `$1`, `$2` สำหรับ capture group) |
| `RegExMode` | text (optional) | `"g"` = ทั้งหมด, `"i"` = ไม่สนตัวพิมพ์ (default: ตัวแรก) |

#### ตัวอย่าง

```
// แทนที่ตัวเลขทั้งหมดด้วย X
ThepRegExReplace("abc 123 def 456", "\d+", "X", "g")
// ผลลัพธ์: "abc X def X"

// ใช้ capture group สลับชื่อ
ThepRegExReplace("John Smith", "(\w+) (\w+)", "$2, $1")
// ผลลัพธ์: "Smith, John"

// ลบ HTML tags
ThepRegExReplace("<b>Hello</b> <i>World</i>", "<[^>]+>", "", "g")
// ผลลัพธ์: "Hello World"
```

---

### ThepRegExMatchCount

นับจำนวน match ของ Regex pattern ในข้อความ

```
ThepRegExMatchCount(OriginalText as text, RegExPattern as text, optional RegExMode as text) as number
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalText` | text | ข้อความต้นฉบับ |
| `RegExPattern` | text | Regex pattern |
| `RegExMode` | text (optional) | `"i"` = case insensitive (นับทั้งหมดเสมอ) |

#### ตัวอย่าง

```
ThepRegExMatchCount("abc 123 def 456", "\d+")
// ผลลัพธ์: 2

ThepRegExMatchCount("Hello hello HELLO", "hello", "i")
// ผลลัพธ์: 3

// ใช้ใน Add Column — นับจำนวน email ในข้อความ
= Table.AddColumn(MyTable, "EmailCount", each ThepRegExMatchCount([Text], "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+"), Int64.Type)
```

---

## Date (จัดการวันที่)

### ThepDatefromText

แปลงข้อความตัวเลข 8 หลักเป็นวันที่ ตาม format ที่กำหนด (ดึงเฉพาะตัวเลข ตัวคั่นอะไรก็ได้)

```
ThepDatefromText(DateText as text, DateFormat as text, optional offsetYear as number) as date
```

| Parameter | Type | Description |
|---|---|---|
| `DateText` | text | ข้อความวันที่ |
| `DateFormat` | text | `"yyyymmdd"`, `"ddmmyyyy"`, หรือ `"mmddyyyy"` |
| `offsetYear` | number (optional) | ปรับปี เช่น `-543` สำหรับ พ.ศ.→ค.ศ. (default: `0`) |

#### ตัวอย่าง

```
ThepDatefromText("20240115", "yyyymmdd")
// ผลลัพธ์: #date(2024, 1, 15)

ThepDatefromText("15/01/2024", "ddmmyyyy")
// ผลลัพธ์: #date(2024, 1, 15)

// แปลง พ.ศ.→ค.ศ.
ThepDatefromText("15-01-2567", "ddmmyyyy", -543)
// ผลลัพธ์: #date(2024, 1, 15)
```

### Tips

- ฟังก์ชันดึงเฉพาะตัวเลขออกมา ดังนั้นตัวคั่นจะเป็น `/`, `-`, `.` หรืออะไรก็ได้ครับ
- ถ้าต้องการแปลงวันที่ที่มี format หลากหลายปนกัน ใช้ **ThepConvertDate** (Premium) แทนครับ

---

### ThepGenDateTableFromDate

สร้างตารางวันที่ต่อเนื่องจาก start date ถึง end date

```
ThepGenDateTableFromDate(StartDate as date, EndDate as date) as table
```

| Parameter | Type | Description |
|---|---|---|
| `StartDate` | date | วันที่เริ่มต้น |
| `EndDate` | date | วันที่สิ้นสุด |

#### ตัวอย่าง

```
ThepGenDateTableFromDate(#date(2024,1,1), #date(2024,12,31))
// ผลลัพธ์: ตาราง 366 แถว (คอลัมน์ "Date") ตั้งแต่ 1 ม.ค. ถึง 31 ธ.ค. 2024
```

### Tips

- ผลลัพธ์มีคอลัมน์เดียวชื่อ `Date` — ต่อยอดด้วย Add Column ได้เลยครับ
- เหมาะสำหรับสร้าง Date Dimension Table ใน Power BI

---

### ThepGenDateTableFromText

เหมือน ThepGenDateTableFromDate แต่รับ input เป็นข้อความ format `yyyymmdd`

```
ThepGenDateTableFromText(StartDateText as text, EndDateText as text) as table
```

| Parameter | Type | Description |
|---|---|---|
| `StartDateText` | text | วันที่เริ่มต้น format `yyyymmdd` |
| `EndDateText` | text | วันที่สิ้นสุด format `yyyymmdd` |

#### ตัวอย่าง

```
ThepGenDateTableFromText("20240101", "20241231")
// ผลลัพธ์: ตาราง 366 แถว (คอลัมน์ "Date")
```

---

### ThepNETWORKDAYS

นับวันทำการระหว่าง 2 วันที่ เหมือน `NETWORKDAYS.INTL` ใน Excel กำหนดวันหยุดสัปดาห์และวันหยุดพิเศษได้ครับ

```
ThepNETWORKDAYS(startDate as date, endDate as date, optional weekendPattern as text, optional holidays as nullable list) as number
```

| Parameter | Type | Description |
|---|---|---|
| `startDate` | date | วันที่เริ่ม |
| `endDate` | date | วันที่สิ้นสุด |
| `weekendPattern` | text (optional) | 7 ตัวอักษร `0`/`1` แทน จ-อา, `1`=หยุด (default: `"0000011"` = เสาร์-อาทิตย์) |
| `holidays` | list (optional) | list ของวันหยุดพิเศษ |

#### ตัวอย่าง

```
// นับวันทำการ (เสาร์-อาทิตย์หยุด)
ThepNETWORKDAYS(#date(2024,1,1), #date(2024,1,31))
// ผลลัพธ์: 23

// หยุดวันศุกร์-เสาร์ (เช่นประเทศในตะวันออกกลาง)
ThepNETWORKDAYS(#date(2024,1,1), #date(2024,1,31), "0000110")

// มีวันหยุดพิเศษ
ThepNETWORKDAYS(
    #date(2024,1,1),
    #date(2024,1,31),
    null,
    {#date(2024,1,1), #date(2024,1,16)}
)
// ผลลัพธ์: 21 (ลบ 2 วันหยุด)
```

### weekendPattern ที่ใช้บ่อย

| Pattern | ความหมาย |
|---|---|
| `"0000011"` | เสาร์-อาทิตย์ (default) |
| `"0000001"` | อาทิตย์อย่างเดียว |
| `"0000110"` | ศุกร์-เสาร์ |
| `"1000001"` | จันทร์-อาทิตย์ |

---

## Table Operations (จัดการตาราง)

### ThepGetColumnName

ดึงชื่อคอลัมน์จากลำดับที่ (1-based) ใช้เมื่อต้องการอ้างอิงคอลัมน์ด้วยตำแหน่งแทนชื่อ

```
ThepGetColumnName(TableName as table, ColNumber as number) as text
```

| Parameter | Type | Description |
|---|---|---|
| `TableName` | table | ตาราง |
| `ColNumber` | number | ลำดับคอลัมน์ (เริ่มจาก 1) |

#### ตัวอย่าง

```
ThepGetColumnName(MyTable, 1)
// ผลลัพธ์: "Name" (ชื่อคอลัมน์แรก)

ThepGetColumnName(MyTable, 3)
// ผลลัพธ์: "Amount" (ชื่อคอลัมน์ที่ 3)
```

---

### ThepGetMultipleListItem

ดึงหลาย item จาก list ด้วย index (0-based) รองรับ index ติดลบ (นับจากท้าย)

```
ThepGetMultipleListItem(OriginalList as list, PosIndex as list) as list
```

| Parameter | Type | Description |
|---|---|---|
| `OriginalList` | list | list ต้นฉบับ |
| `PosIndex` | list | list ของ index ที่ต้องการ (0-based, ติดลบได้) |

#### ตัวอย่าง

```
ThepGetMultipleListItem({"A","B","C","D","E"}, {0, 2, 4})
// ผลลัพธ์: {"A", "C", "E"}

// ใช้ index ติดลบ (นับจากท้าย)
ThepGetMultipleListItem({"A","B","C","D","E"}, {-1, -2})
// ผลลัพธ์: {"E", "D"}
```

---

### ThepRenameColumn

เปลี่ยนชื่อคอลัมน์ด้วยเลขลำดับ (1-based) ไม่ต้องรู้ชื่อเดิม ส่งทีละตัวหรือหลายตัวก็ได้ครับ

```
ThepRenameColumn(TableName as table, ColNumber as any, NewName as any) as table
```

| Parameter | Type | Description |
|---|---|---|
| `TableName` | table | ตาราง |
| `ColNumber` | number หรือ list | ลำดับคอลัมน์ (1-based) |
| `NewName` | text หรือ list | ชื่อใหม่ |

#### ตัวอย่าง

```
// เปลี่ยนชื่อคอลัมน์ที่ 1
ThepRenameColumn(MyTable, 1, "ProductName")

// เปลี่ยนหลายคอลัมน์พร้อมกัน
ThepRenameColumn(MyTable, {1, 3, 5}, {"Name", "Amount", "Date"})
```

---

### ThepReplaceAllError

แทนที่ error values ทุกคอลัมน์ในตารางพร้อมกัน ไม่ต้องทำทีละคอลัมน์ครับ

```
ThepReplaceAllError(TableName as table, optional ReplaceWith as text) as table
```

| Parameter | Type | Description |
|---|---|---|
| `TableName` | table | ตาราง |
| `ReplaceWith` | text (optional) | ค่าที่ใช้แทน error (default: `null`) |

#### ตัวอย่าง

```
// แทน error ทุกคอลัมน์ด้วย null
ThepReplaceAllError(MyTable)

// แทน error ด้วย "N/A"
ThepReplaceAllError(MyTable, "N/A")
```

---

### ThepOneHot

One-Hot Encoding — แปลงค่าในคอลัมน์เป็นคอลัมน์ใหม่ แต่ละค่าเป็น 1 คอลัมน์ เหมาะกับงาน data prep สำหรับ Machine Learning ครับ

```
ThepOneHot(TableName as table, TargetColumnName as text) as table
```

| Parameter | Type | Description |
|---|---|---|
| `TableName` | table | ตาราง |
| `TargetColumnName` | text | ชื่อคอลัมน์ที่ต้องการ encode |

#### ตัวอย่าง

สมมติมีตาราง:
| Name | Color |
|---|---|
| A | Red |
| B | Blue |
| C | Red |

```
ThepOneHot(MyTable, "Color")
```

ผลลัพธ์:
| Name | Color | Red | Blue |
|---|---|---|---|
| A | Red | 1 | 0 |
| B | Blue | 0 | 1 |
| C | Red | 1 | 0 |

---

### ThepOpenAI

เรียก OpenAI Chat API จาก Power Query (เวอร์ชัน Free)

```
ThepOpenAI(apiKey as text, userPrompt as text, optional systemPrompt as nullable text, optional model as nullable text) as text
```

| Parameter | Type | Description |
|---|---|---|
| `apiKey` | text | OpenAI API Key |
| `userPrompt` | text | ข้อความที่ส่งให้ AI |
| `systemPrompt` | text (optional) | System prompt |
| `model` | text (optional) | ชื่อ model (default: `"gpt-4o-mini"`) |

> **แนะนำ:** ใช้ **ThepLLM** (Premium) แทนครับ รองรับหลาย provider, มี error handling ที่ดีกว่า, และ default เป็น OpenRouter ที่เข้าถึง 200+ models ด้วย key เดียว

#### ตัวอย่าง

```
ThepOpenAI(OpenAI_Key, "สรุปข้อความนี้: " & [Text])
```
