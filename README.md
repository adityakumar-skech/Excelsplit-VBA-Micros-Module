# excel-split
divide excel using number of rows vba script 
##################################################
You can do it this way with a helper column:

1st -- insert a column at the left (i.e column A).

2nd -- fill down the new column in a series from 1 to 3000 (or however many rows you have).

3rd -- in the two cells below the end of the series of column A, put in 100.5 and 200.5. Highlight these cells, then use Ctrl + Shift + Down Arrow and choose Fill Series (leaving the Step Value at 100)

4th -- Ascending Sort Column A.

5th -- Delete Column A. You'll have a blank row every 100 rows, with whatever else left over.
