### Examples for some lambdas included in [`my_excel_lambda_functions.txt`](https://raw.githubusercontent.com/alekrutkowski/ExcelLambdaTools/main/my_excel_lambda_functions.txt)

If you download the data from https://ec.europa.eu/eurostat/api/dissemination/sdmx/2.1/data/NAMA_10_PC/.CP_EUR_HAB+CP_NAC_HAB.B1GQ+P31_S14.BE+BG+CZ+DE?format=TSV&startPeriod=2020&endPeriod=2022
you will get this:

|         | A     | B           | C        | D    | E        | F        | G        |
|---------|-------|-------------|----------|------|----------|----------|----------|
| **1**   | freq  | unit        | na_item  | geo  | 2020     | 2021     | 2022     |
| **2**   | A     | CP_EUR_HAB  | B1GQ     | BE   | 39830 p  | 43350 p  | 46990 p  |
| **3**   | A     | CP_EUR_HAB  | B1GQ     | BG   | 8890     | 10330    | 12400 p  |
| **4**   | A     | CP_EUR_HAB  | B1GQ     | CZ   | 20170    | 22270    | 25850    |
| **5**   | A     | CP_EUR_HAB  | B1GQ     | DE   | 40930 p  | 43480 p  | 46260 p  |
| **6**   | A     | CP_EUR_HAB  | P31_S14  | BE   | 19250 p  | 20740 p  | 23240 p  |
| **7**   | A     | CP_EUR_HAB  | P31_S14  | BG   | 5150     | 5980     | 7370 p   |
| **8**   | A     | CP_EUR_HAB  | P31_S14  | CZ   | 8960     | 9900     | 11820    |
| **9**   | A     | CP_EUR_HAB  | P31_S14  | DE   | 19890 p  | 20800 p  | 22930 p  |
| **10**  | A     | CP_NAC_HAB  | B1GQ     | BE   | 39830 p  | 43350 p  | 46990 p  |
| **11**  | A     | CP_NAC_HAB  | B1GQ     | BG   | 17390    | 20210    | 24250 p  |
| **12**  | A     | CP_NAC_HAB  | B1GQ     | CZ   | 533560   | 571050   | 634910   |
| **13**  | A     | CP_NAC_HAB  | B1GQ     | DE   | 40930 p  | 43480 p  | 46260 p  |
| **14**  | A     | CP_NAC_HAB  | P31_S14  | BE   | 19250 p  | 20740 p  | 23240 p  |
| **15**  | A     | CP_NAC_HAB  | P31_S14  | BG   | 10060    | 11700    | 14410 p  |
| **16**  | A     | CP_NAC_HAB  | P31_S14  | CZ   | 237040   | 253890   | 290420   |
| **17**  | A     | CP_NAC_HAB  | P31_S14  | DE   | 19890 p  | 20800 p  | 22930 p  |

If you use the function **`=reshapeToLong(A1:G17,{"unit","na_item","geo"},{2020,2021,2022},"val","year")`** in any Excel cell that has sufficiently many
empty cells below and to the right (to avoid the [#SPILL! error](https://support.microsoft.com/en-us/office/how-to-correct-a-spill-error-ffe0f555-b479-4a17-a6e2-ef9cc9ad4023#:~:text=This%20error%20occurs%20when%20the,the%20obstructing%20cell(s).)) you will get this:

| val      | year  | unit        | na_item  | geo  |
|----------|-------|-------------|----------|------|
| 39830 p  | 2020  | CP_EUR_HAB  | B1GQ     | BE   |
| 8890     | 2020  | CP_EUR_HAB  | B1GQ     | BG   |
| 20170    | 2020  | CP_EUR_HAB  | B1GQ     | CZ   |
| 40930 p  | 2020  | CP_EUR_HAB  | B1GQ     | DE   |
| 19250 p  | 2020  | CP_EUR_HAB  | P31_S14  | BE   |
| 5150     | 2020  | CP_EUR_HAB  | P31_S14  | BG   |
| 8960     | 2020  | CP_EUR_HAB  | P31_S14  | CZ   |
| 19890 p  | 2020  | CP_EUR_HAB  | P31_S14  | DE   |
| 39830 p  | 2020  | CP_NAC_HAB  | B1GQ     | BE   |
| 17390    | 2020  | CP_NAC_HAB  | B1GQ     | BG   |
| 533560   | 2020  | CP_NAC_HAB  | B1GQ     | CZ   |
| 40930 p  | 2020  | CP_NAC_HAB  | B1GQ     | DE   |
| 19250 p  | 2020  | CP_NAC_HAB  | P31_S14  | BE   |
| 10060    | 2020  | CP_NAC_HAB  | P31_S14  | BG   |
| 237040   | 2020  | CP_NAC_HAB  | P31_S14  | CZ   |
| 19890 p  | 2020  | CP_NAC_HAB  | P31_S14  | DE   |
| 43350 p  | 2021  | CP_EUR_HAB  | B1GQ     | BE   |
| 10330    | 2021  | CP_EUR_HAB  | B1GQ     | BG   |
| 22270    | 2021  | CP_EUR_HAB  | B1GQ     | CZ   |
| 43480 p  | 2021  | CP_EUR_HAB  | B1GQ     | DE   |
| 20740 p  | 2021  | CP_EUR_HAB  | P31_S14  | BE   |
| 5980     | 2021  | CP_EUR_HAB  | P31_S14  | BG   |
| 9900     | 2021  | CP_EUR_HAB  | P31_S14  | CZ   |
| 20800 p  | 2021  | CP_EUR_HAB  | P31_S14  | DE   |
| 43350 p  | 2021  | CP_NAC_HAB  | B1GQ     | BE   |
| 20210    | 2021  | CP_NAC_HAB  | B1GQ     | BG   |
| 571050   | 2021  | CP_NAC_HAB  | B1GQ     | CZ   |
| 43480 p  | 2021  | CP_NAC_HAB  | B1GQ     | DE   |
| 20740 p  | 2021  | CP_NAC_HAB  | P31_S14  | BE   |
| 11700    | 2021  | CP_NAC_HAB  | P31_S14  | BG   |
| 253890   | 2021  | CP_NAC_HAB  | P31_S14  | CZ   |
| 20800 p  | 2021  | CP_NAC_HAB  | P31_S14  | DE   |
| 46990 p  | 2022  | CP_EUR_HAB  | B1GQ     | BE   |
| 12400 p  | 2022  | CP_EUR_HAB  | B1GQ     | BG   |
| 25850    | 2022  | CP_EUR_HAB  | B1GQ     | CZ   |
| 46260 p  | 2022  | CP_EUR_HAB  | B1GQ     | DE   |
| 23240 p  | 2022  | CP_EUR_HAB  | P31_S14  | BE   |
| 7370 p   | 2022  | CP_EUR_HAB  | P31_S14  | BG   |
| 11820    | 2022  | CP_EUR_HAB  | P31_S14  | CZ   |
| 22930 p  | 2022  | CP_EUR_HAB  | P31_S14  | DE   |
| 46990 p  | 2022  | CP_NAC_HAB  | B1GQ     | BE   |
| 24250 p  | 2022  | CP_NAC_HAB  | B1GQ     | BG   |
| 634910   | 2022  | CP_NAC_HAB  | B1GQ     | CZ   |
| 46260 p  | 2022  | CP_NAC_HAB  | B1GQ     | DE   |
| 23240 p  | 2022  | CP_NAC_HAB  | P31_S14  | BE   |
| 14410 p  | 2022  | CP_NAC_HAB  | P31_S14  | BG   |
| 290420   | 2022  | CP_NAC_HAB  | P31_S14  | CZ   |
| 22930 p  | 2022  | CP_NAC_HAB  | P31_S14  | DE   |

