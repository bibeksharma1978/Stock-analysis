# Stock Analysis With Excel VBA

## Purpose and Background
The purpose of this project was to see if the refactor code was running fast 

## Results
the refractor code was very fast and efficient.
### Analysis
for the refactoring the code, I i found out that if the output arrays are not placed accordingly it is not going to do the output as needed, here is eg of what i did wrong

'1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
     Dim tickerEndingPrices(12) As Single
     Dim tickerStartingPrices(12) As Single
     Dim tickerVolumes(12) As Long

this was later corrected and it workedl, the corrections is as following

    '1a) Create a ticker Index         
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    

## Summary
### Pros and Cons of Refactoring Code
Refactoring the codes are simple to understand and the cons is that its time consuming to find the error 

