# v16 fetch log

## course_x_po (grid-tUkDsBwjVg)
- Filter: `[Program Abbreviated] = "MSCSIA"`
- filterColumnNames: _progs|_courses, Program Learning Outcome, Course IRMA Map
- Result: totalRows=55, hasMore=false, single page ✓
- Payload: ~minimal (3 cols/row)

## course_x_cct (grid-bdbEJT9Kuq)
- Filter: `[Program Abbreviated] = "MSCSIA"` ← NEW pattern
- filterColumnNames: Program to Course Pairing, Cross Cutting Theme, Aligned
- Expected: 44 rows, single page

## comp_x_po (grid-NkK5JDeDhF)
- Filter: `[Program Abbreviation] = "MSCSIA"` ← different column name
- filterColumnNames: Program Course Competency, Program Outcome, Course Comp IRMA Map to PO
- Expected: 200 rows, 2 pages

## comp_x_cct (grid-6xLuZLsQ_t)
- Filter: `[Program Abbreviation] = "MSCSIA"`
- filterColumnNames: Program Course Competency, Cross-Cutting Theme, Aligned
- Expected: 160 rows, 2 pages
