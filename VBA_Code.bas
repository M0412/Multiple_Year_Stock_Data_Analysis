{\rtf1\ansi\ansicpg1252\cocoartf2709
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;\red251\green0\blue44;\red253\green139\blue9;\red0\green0\blue0;
\red36\green255\blue255;\red251\green0\blue255;\red255\green255\blue11;\red251\green0\blue7;\red190\green255\blue255;
\red61\green57\blue1;\red79\green50\blue1;\red80\green39\blue1;\red46\green0\blue141;\red41\green6\blue19;
\red94\green1\blue38;\red196\green255\blue139;\red34\green255\blue6;\red187\green187\blue187;\red27\green128\blue255;
}
{\*\expandedcolortbl;;\cssrgb\c100000\c1320\c22539;\cssrgb\c100000\c61456\c0;\cssrgb\c0\c0\c0;
\cssrgb\c4983\c100000\c100000;\cssrgb\c100000\c7248\c100000;\cssrgb\c100000\c100000\c0;\cssrgb\c100000\c12195\c0;\cssrgb\c78354\c100000\c100000;
\cssrgb\c30438\c28381\c0;\cssrgb\c38621\c25545\c0;\cssrgb\c39024\c20428\c0;\cssrgb\c24324\c6455\c62064;\cssrgb\c21766\c3110\c9785;
\cssrgb\c44903\c5525\c19561;\cssrgb\c80504\c100000\c61376;\cssrgb\c0\c100000\c0;\cssrgb\c78156\c78156\c78156;\cssrgb\c11095\c58865\c100000;
}
\margl1440\margr1440\vieww28300\viewh14840\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf2 Attribute VB_Name = "Module1"\cf0 \
\
\cf3 Sub \cf4 Stock_Date()\cf0 \
\cf5 \
            ' \'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97-\
            ' LOOP THROUGH ALL THE SHEETS\
            ' \'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97-\cf0 \
      \
            \cf6  For \cf4 Each ws In Worksheets\cf0 \
          \
                         \cf5 ' CREATING COLUMN HEADERS\cf0 \
                          ws.Cells(1, 9).Value = "Ticker"\
                          ws.Cells(1, 10).Value = "Yearly Change"\
                          ws.Cells(1, 11).Value = "Percent Change"\
                          ws.Cells(1, 12).Value = "Total Stock Volume"\
                          ws.Cells(2, 15).Value = "Greatest % Increase"\
                          ws.Cells(3, 15).Value = "Greatest % Decrease"\
                          ws.Cells(4, 15).Value = "Greatest Total Volume"\
                          ws.Cells(1, 16).Value = "Ticker"\
                          ws.Cells(1, 17).Value = "Value"\
      \
                       \cf5   ' \'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\
                         ' LOOP THROUGH ALL TICKERS\
                         ' \'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\cf0 \
           \
          \cf5                ' DETERMINE THE LAST ROW\cf0 \
                         LR = ws.Cells(Rows.Count, 1).End(xlUp).Row\
             \
                         \cf5 ' KEEP TRACK OF THE LOCATION FOR EACH TICKER NAME IN THE SUMMARY TABLE\cf0 \
                         Summary_TableRow = 2\
           \
                       \cf5  ' SETTING A VARIABLE\cf0 \
                         StartOfTicker = 2\
           \
                                   \cf7   For\cf4  i = 2 To LR\cf0 \
           \
                                             \cf5  ' CHECK TO SEE IF WE ARE SILL IN THE SAME TICKER, WE ARE NOT, THEN PRINT RESULTS\cf0 \
                                             \cf8  If\cf0  ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value \cf8 Then\cf0 \
            \
                                                         \cf5 ' SET THE TICKER NAME\cf0 \
                                                          Ticker_Name = ws.Cells(i, 1).Value\
                           \
                                                         \cf5 ' SET THE OPEN DATE\cf0 \
                                                          Open_Date = ws.Cells(StartOfTicker, 3).Value\
                         \
                                                        \cf5  ' SET THE CLOSE DATE\cf0 \
                                                         Close_Date = ws.Cells(i, 6).Value\
                         \
                                                        \cf5  ' ADD TO THE STOCK VOLUME\cf0 \
                                                          Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value\
                           \
                                                        \cf5  ' PRINT TICKER NAME IN THE SUMMARY TABLE\cf0 \
                                                         ws.Range("I" & Summary_TableRow).Value = Ticker_Name\
                           \
                                                        \cf5  ' PRINT YEARLY CHANGE IN THE SUMMARY TABLE\cf0 \
                                                         ws.Range("J" & Summary_TableRow).Value = Close_Date - Open_Date\
                           \
                                                         \cf5 ' HIGHLIGHT YEARLY CHANGE RESULTS BASED ON VALUE\cf0 \
                                                       \cf9  \cf10  \cf11 If \cf0 ws.Range("J" & Summary_TableRow).Value > 0\cf9  \cf11 Then\cf0 \
                                                                 ws.Range("J" & Summary_TableRow).Interior.ColorIndex = 4\
                                  \
                                                         \cf12 ElseIf\cf0  ws.Range("J" & Summary_TableRow).Value < 0 \cf12 Then\cf0 \
                                                                 ws.Range("J" & Summary_TableRow).Interior.ColorIndex = 3\
                                  \
                                                         \cf12 Else\cf0 \
                                                                 ws.Range("J" & Summary_TableRow).Interior.ColorIndex = 6\
\
                                                        \cf12  End If\cf0 \
                           \
                                                         \cf5 ' PRINT PERCENT CHANGE IN THE SUMMARY TABLE\cf0 \
                                                         ws.Range("K" & Summary_TableRow).Value = ws.Range("J" & Summary_TableRow).Value / Open_Date\
                           \
                                                      \cf5   ' HIGHLIGHT PERCENT CHANGE RESULTS BASED ON VALUES\cf0 \
                                                       \cf13  If\cf0  ws.Range("K" & Summary_TableRow).Value > 0 \cf13 Then\cf0 \
                                                               ws.Range("K" & Summary_TableRow).Interior.ColorIndex = 4\
                                  \
                                                        \cf13 ElseIf\cf0  ws.Range("K" & Summary_TableRow).Value < 0 \cf13 Then\cf0 \
                                                               ws.Range("K" & Summary_TableRow).Interior.ColorIndex = 3\
                                  \
                                                        \cf13 Else\cf0 \
                                                            ws.Range("K" & Summary_TableRow).Interior.ColorIndex = 6\
\
                                                        \cf13 End If\cf0 \
                           \
                                                       \cf5  ' PRINT STOCK VOLUME IN THE SUMMARY TABLE\cf0 \
                                                        ws.Range("L" & Summary_TableRow).Value = Stock_Volume\
                           \
                                                       \cf5  ' ADD ONE TO THE SUMMARY TABLE ROW\cf0 \
                                                        Summary_TableRow = Summary_TableRow + 1\
                           \
                                                       \cf5  ' RESET THE STOCK VOLUME\cf0 \
                                                        Stock_Volume = 0\
                           \
                                                        \cf5 ' START OF THE NEXT TICKER\cf0 \
                                                        StartOfTicker = i + 1\
           \
                                           \cf5  ' IF THE CELL IMMEDIATELY FOLLOWING A ROW IS THE SAME TICKER\cf0 \
                                            \cf8 Else\cf0 \
                    \
                                                       \cf5  ' ADD TO THE STOCK VOLUME\cf0 \
                                                        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value\
                          \
                                            \cf8 End If\cf0 \
           \
                                \cf7  Next\cf4  i\cf0 \
        \
                                \cf5  ' FIND THE GREATEST TOTAL VOLUME\cf0 \
                                  TotalVolumeLastRow = ws.Cells(Rows.Count, 12).End(xlUp).Row\
           \
                                 Max_Total_Volume = Application.WorksheetFunction.Max(ws.Range("L:L"))\
            \
                               \cf14   \cf15 For\cf4  j = 2\cf14  \cf0 To TotalVolumeLastRow\
                 \
                                        \cf16  If \cf0 ws.Cells(j, 12) = Max_Total_Volume \cf16 Then\cf0 \
                             \
                                                    ws.Cells(4, 16).Value = ws.Cells(j, 9)\
                                                    ws.Cells(4, 17).Value = Max_Total_Volume\
                                 \
                                         \cf16 End If\cf0 \
                \
                                \cf15 Next\cf4  j\cf0 \
                \
                              \cf5  ' FIND THE MAX PERCENT INCREASE AND DECREASE\cf0 \
                               PercentChangeLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row\
             \
                               Max_Percent_Increase = Application.WorksheetFunction.Max(ws.Range("K:K"))\
                               Max_Percent_Decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))\
             \
                              \cf17  For\cf18  \cf0 l = 2 To PercentChangeLastRow\
                 \
                                      \cf19  If\cf0  ws.Cells(l, 11) = Max_Percent_Increase \cf19 Then\cf0 \
                                 \
                                                 ws.Cells(2, 16).Value = ws.Cells(l, 9)\
                                                 ws.Cells(2, 17).Value = Max_Percent_Increase\
                     \
                                       \cf19 ElseIf\cf0  ws.Cells(l, 11) = Max_Percent_Decrease \cf19 Then\cf0 \
                                 \
                                                 ws.Cells(3, 16).Value = ws.Cells(l, 9)\
                                                 ws.Cells(3, 17).Value = Max_Percent_Decrease\
                    \
                                     \cf19 End If\cf0 \
               \
                             \cf17 Next\cf0  l\
          \
           \cf6  Next\cf0 \
    \
\cf3 End Sub\cf0 \
}
