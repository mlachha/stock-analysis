# STOCK ANALYSIS 

### OVERVIEW: VBA Stock Analysis Project

A stock is an investment that represents a share, or partial ownership, of a company. Stocks are one of the best ways to build wealth. 
In this project we have a dataset in excel that contains a little over 3000 daily records for daily volums exchanged per stock, for the years 2017 and 2018. 

### Purpose

In order to figure out which stocks perform best were going to look at how each stock performed over the year, by analysing the total volum of exchange and the yearly return.

## Analysis and Challenges

### Challenges:

The data set is made up of exactly 3013 example which is a tiny percentage of the real number of exchanges made over the year. based on how limited it is, the analysis cant be considered to be conclusive and should serve for familiarisation purposes only.

### Analysis:

 the client is intersted in a specific stock 'DADQ' with the ticker 'DQ'.
 for that in vba we are going to build a script that calculates the total yearly traded volum based on the daily volums for the ticker 'DQ'.
 - code looks like:
 
 ```vba
 Sub DQAnalysis()
Worksheets("DQAnalysis").Activate
Range("A1").Value = "DAQO (Ticker: DQ)"
'Create a header row
Cells(3, 1).Value = "Year"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"  
    Worksheets("2018").Activate
    totalVolume = 0
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
rowStart = 2
'DELETE: rowEnd = 3013
'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

For i = rowStart To rowEnd
    'increase totalVolume
     If Cells(i, 1).Value = "DQ" Then
        totalVolume = totalVolume + Cells(i, 8).Value
        End If
    
     If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
            startingPrice = Cells(i, 6).Value
        End If

        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
            endingPrice = Cells(i, 6).Value
        End If

Next i

    Worksheets("DQAnalysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1
End Sub
 
 ```
 
 
