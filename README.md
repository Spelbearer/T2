T2

This repository contains a PyQt application for converting tabular data into

KML files. Numerical grouping ranges use Jenks natural breaks with
a fallback to equal intervals when data variety is low or when Jenks
does not provide the requested number of boundaries. Zero values are
considered when calculating ranges. The interface now allows selecting
any number of groups up to twenty without the value being automatically
reduced. Only the first twenty rows of the loaded file are shown in the
preview table. Data can be filtered using a pandas-style expression
entered in the **Filter formula** field to generate multiple layers from a
single source.
