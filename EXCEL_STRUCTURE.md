# Excel Sheet Structure - Vocabulary Tracker

## Column Definitions

| Column | Name | Description |
|--------|------|-------------|
| A | Order | Sequential order/number of the word |
| B | Topic/Session | Name of the reading/listening session where words are from |
| C | Flag | Learning status indicator |
| D | Word | The vocabulary word (English) |
| E | Part of Speech | Grammar category (noun, verb, adjective, etc.) |
| F | Pronunciation | Phonetic pronunciation |
| G | Meaning | Vietnamese translation/meaning |
| H | Example Sentence | Sentence using the word in context |
| I | Synonyms | Alternative words with similar meaning |
| J | Day of Week | Day when the word was caught (1-8) |
| K | Date | Date when the word was caught (m/d/yyyy format) |

## Flag Values

| Flag | Meaning |
|------|---------|
| `N` or `n` | **New word** - Don't know this word, need to learn it |
| `Y` or `y` | **Known important** - Know it thoroughly but it's important |
| `?` | **Forgotten** - Learned before but don't remember exactly, need to re-learn |
| `ok` | **Recently learned** - Didn't know at first but learned it. Still need to learn by heart |

## Data Structure

The Excel sheet follows a hierarchical structure:

```
Row 1: [Order] [TOPIC NAME] [Flag] [Word] [POS] [Pronunciation] [Meaning] [Sentence] [Synonyms] [Day] [Date]
Row 2: [Order] [empty]      [Flag] [Word] [POS] [Pronunciation] [Meaning] [Sentence] [Synonyms] [Day] [Date]
Row 3: [Order] [empty]      [Flag] [Word] [POS] [Pronunciation] [Meaning] [Sentence] [Synonyms] [Day] [Date]
Row 4: [Order] [TOPIC NAME] [Flag] [Word] [POS] [Pronunciation] [Meaning] [Sentence] [Synonyms] [Day] [Date]
Row 5: [Order] [empty]      [Flag] [Word] [POS] [Pronunciation] [Meaning] [Sentence] [Synonyms] [Day] [Date]
```

### Rules:
1. When **Column B has a value** (topic name), it marks the **start of a new topic/session**
2. All rows below that row (where Column B is empty) **belong to that topic**
3. This continues until another row with a topic name in Column B appears
4. **Column D** always contains the vocabulary word for each row

## Example

| A | B | C | D | E | F | G | H | I | J | K |
|---|---|---|---|---|---|---|---|---|---|---|
| 1 | TOEIC Reading Part 7 | N | abandon | verb | /əˈbændən/ | từ bỏ | He abandoned his car. | desert, leave | 5 | 2/9/2026 |
| 2 | | Y | significant | adj | /sɪɡˈnɪfɪkənt/ | quan trọng | A significant change occurred. | important, notable | 5 | 2/9/2026 |
| 3 | | ? | enhance | verb | /ɪnˈhæns/ | cải thiện | We need to enhance quality. | improve, boost | 5 | 2/9/2026 |
| 4 | Listening Practice 1 | ok | furthermore | adv | /ˈfɜːrðərmɔːr/ | hơn nữa | Furthermore, we must act now. | moreover, additionally | 6 | 2/10/2026 |
| 5 | | N | evaluate | verb | /ɪˈvæljueɪt/ | đánh giá | They will evaluate the results. | assess, judge | 6 | 2/10/2026 |

In this example:
- Rows 1-3 belong to topic "TOEIC Reading Part 7"
- Rows 4-5 belong to topic "Listening Practice 1"
