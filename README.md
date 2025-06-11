T2

This repository contains a PyQt application for converting tabular data into

KML files. Numerical grouping ranges are determined using Jenks natural breaks
whenever possible, falling back to equal intervals when data variety is low.
Zero values are included in the range calculation. The number of groups can be
configured from three to five via the interface.
=======
KML files. Numerical grouping ranges are now calculated using equal intervals
from the minimum (including any zero values) to the maximum value. The number
of groups can be configured from three to five via the interface.

