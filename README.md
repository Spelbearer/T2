

This repository contains a PyQt application for converting tabular data into

KML files. Numerical grouping ranges use Jenks natural breaks with
a fallback to equal intervals when data variety is low or when Jenks
does not provide the requested number of boundaries. Zero values are
considered when calculating ranges. The interface now allows selecting
any number of groups up to twenty without the value being automatically
reduced. Only the first twenty rows of the loaded file are shown in the

preview table. Data may be filtered with a simple expression in the
**Filter formula** field, for example `City=London and Value>5`.
Numeric fields in the expression are converted automatically so
comparisons such as `<` and `>` function correctly. The program
translates the formula into a pandas query so multiple layers can be
generated from a single source.


preview table. Data may be filtered with a simple expression in the
**Filter formula** field, for example `City=London and Value>5`.
Numeric fields are automatically detected so comparison operators like
`<` and `>` work as expected. The program converts this formula into a
pandas query so multiple layers can be generated from a single source.


preview table. Data may be filtered with a simple expression in the
**Filter formula** field, for example `City=London and Value>5`. The
program converts this formula into a pandas query so multiple layers can
be generated from a single source.

preview table. Data can be filtered using a pandas-style expression
entered in the **Filter formula** field to generate multiple layers from a
single source.
