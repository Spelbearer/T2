T2

This repository contains a PyQt application for converting tabular data into

KML files. Numerical grouping ranges use Jenks natural breaks with
a fallback to equal intervals when data variety is low or when Jenks
does not provide the requested number of boundaries. Zero values are
considered when calculating ranges. The interface now allows selecting
any number of groups up to twenty without the value being automatically
reduced. Only the first twenty rows of the loaded file are shown in the
preview table. Data may be filtered with a simple expression in the
**Filter formula** field, for example `City=London and Value>5`.
Numeric fields are automatically detected so comparison operators like
`<` and `>` work as expected. The program converts this formula into a
pandas query so multiple layers can be generated from a single source.
