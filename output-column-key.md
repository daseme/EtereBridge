# Output File Column Key

| Column Name   | Source / Origin                                           | Treatment / Transformation                                                                              | Expected Format            |
|---------------|-----------------------------------------------------------|---------------------------------------------------------------------------------------------------------|----------------------------|
| Bill Code     | Textbox180 & Textbox171 from CSV header                   | Combined as "Textbox180:Textbox171"; if one is missing, use the available value                         | Plain text                 |
| Time In       | Derived from CSV “timerange2” split                       | Parsed to a 24‑hour time; converted to Excel time serial; formatted if valid                             | HH:MM:SS                   |
| Time Out      | Derived from CSV “timerange2” split                       | Parsed similarly to Time In                                                                               | HH:MM:SS                   |
| Air Date      | CSV “dateschedule” column (renamed)                       | Cleaned, converted to datetime; verified and reformatted                                               | m/d/yy                     |
| End Date      | Based on Air Date & template formulas                     | Formula-driven: either copied from template or auto-linked to Air Date                                  | m/d/yy (via formula)       |
| Month         | Derived from Air Date and Billing Type                    | If Billing Type = Calendar, same as Air Date; if Broadcast, compute “broadcast month” (first day of month) | mmm-yy                     |
| Priority      | Added during Excel save routine                           | Hardcoded to a default value (4)                                                                         | Numeric (4)                |
| Gross Rate    | CSV “IMPORTO2” (renamed)                                  | Cleaned (remove $/commas), converted to numeric, then re‑formatted as currency                          | $#,##0.00                  |
| Length        | CSV “duration3” (renamed)                                 | Rounded to nearest 15 sec; converted from seconds to HH:MM:SS                                            | HH:MM:SS                   |
| Line / #      | CSV “id_contrattirighe” & “Textbox14” (renamed)           | Converted to numeric and integer form                                                                    | Integer                    |
| Market        | CSV “nome2” (renamed)                                     | Market replacements applied per config                                                                 | Standardized market names  |
| Program       | CSV “airtimep” (renamed)                                  | Passed through as text                                                                                   | Text                       |
| Media         | CSV “bookingcode2” (renamed)                              | Used as-is (identifies media type)                                                                       | Text                       |
| Billing Type  | User input prompt                                         | Added as a new column                                                                                     | "Calendar" or "Broadcast"  |
| Revenue Type  | User input prompt                                         | Added as a new column                                                                                     | Text                     |
| Agency?       | User input prompt                                         | Added as a new column; determines fee application                                                       | "Agency", "Non-Agency", etc.|
| Sales Person  | User input prompt                                         | Added as a new column                                                                                     | Text                       |
| Lang.         | Derived via language detection on “rowdescription”        | Mapped to language code                                                                                   | One of: E, M, T, Hm, SA, V…  |
| Affidavit?    | User input prompt                                         | Added as a new column                                                                                     | "Y" or "N"                 |
| Estimate      | User input prompt                                         | Added as a new column (optional)                                                                          | Text or empty             |
| Contract      | User input prompt                                         | Added as a new column                                                                                     | Text (non‑empty)           |
| Type          | Computed from “Gross Rate”                                | If Gross Rate equals zero → “BNS”, else “COM”                                                             | "BNS" or "COM"             |
| Broker Fees   | Calculated if Agency? is “Agency”                         | Computed as Gross Rate * agency fee rate; formatted as currency                                           | $#,##0.00 (or blank)       |
