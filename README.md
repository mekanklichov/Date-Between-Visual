# Year Range Filter - Power BI Custom Visual

A compact two-dropdown custom visual for Power BI.

## Data field

Add exactly one field to the visual's **Year** data role:

`Dim_Date[Year]`

## Behavior

- Left dropdown: ascending years.
- Right dropdown: descending years.
- Defaults to minimum and maximum available year.
- Changing either dropdown applies a Power BI Advanced Filter to the bound year column.
- The filter is applied through the visual host, so it can affect other visuals on the report page according to normal Power BI filter interactions.
- No title is rendered.
- No bottom chart or extra decoration is rendered.

## Package

Install the Power BI Visuals tooling in an environment with internet access, then run:

`npm install`

`npm run package`

The generated `.pbiviz` will be placed in `dist/`.
