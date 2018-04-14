`calc2xml`
==========

`calc2xml` is a converter that converts (on the command line) an [LibreOffice Calc](https://www.libreoffice.org/discover/calc/) ODS file into a corresponding [XML document](https://www.w3.org/TR/REC-xml/) using an [XSL transformation](https://www.w3.org/TR/1999/REC-xslt-19991116).

Only certain output tables are taken into account for output. The output XML is generated according to a simple scheme.

The first row of the output table must contain unique column identifiers: Cell $A1$ an id for column $A$, $B1$ an id (different from $A1$) for column $B$, aso.

There must not be an empty column up to the last filled column.

A table with, for example, five columns must therefore fill the columns $A$ to $E$ and the cells $A1$ to $E1$ must contain different identifiers.

The cells $A2$, $A3$, … contain data associated with the identifier in the $A1$ cell, the cells $B2$, $B3$, … according to data associated with the identifier in the $B1$ cell, aso. up to the $E2$, … cells containing data associated with the identifier in the $E1$ cell.
