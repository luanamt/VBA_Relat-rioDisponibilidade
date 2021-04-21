Option Explicit

Sub ExportToWord()

Dim WordApp As New Word.Application
Dim doc As Word.Document

Dim rng_CSC_DTOTAL As Word.Range
Dim rng_CSC_DLINK As Word.Range
Dim rng_CSC_DispTot1 As Word.Range
Dim rng_CSC_DispTot2 As Word.Range
Dim rng_CSC_DispTot1p As Word.Range
Dim rng_CSC_DispTot2p As Word.Range

Dim rng_CSC_DispL1 As Word.Range
Dim rng_CSC_DispL2 As Word.Range
Dim rng_CSC_DispL1p As Word.Range
Dim rng_CSC_DispL2p As Word.Range

Dim rng_Grupo1 As Word.Range
Dim rng_Grupo2 As Word.Range

Dim rng_Taxa1 As Word.Range
Dim rng_Taxa2 As Word.Range

Dim rng_CSC_Prob1 As Word.Range
Dim rng_CSC_Prob2 As Word.Range

Dim rng_CC_DTOTAL As Word.Range
Dim rng_CC_DLINK As Word.Range
Dim rng_CC_DispTot1 As Word.Range
Dim rng_CC_DispTot2 As Word.Range
Dim rng_CC_DispTot1p As Word.Range
Dim rng_CC_DispTot2p As Word.Range

Dim rng_CC_DispL1 As Word.Range
Dim rng_CC_DispL2 As Word.Range
Dim rng_CC_DispL1p As Word.Range
Dim rng_CC_DispL2p As Word.Range

Dim rng_CC_Prob1 As Word.Range
Dim rng_CC_Prob2 As Word.Range

Set doc = WordApp.Documents.Open([wordPath].Text, , True)



Set rng_Grupo1 = doc.Bookmarks("Grupo1").Range
rng_Grupo1.Text = Range("Grupo1")
Set rng_Grupo2 = doc.Bookmarks("Grupo2").Range
rng_Grupo2.Text = Range("Grupo2")
Set rng_Taxa1 = doc.Bookmarks("Taxa1").Range
rng_Taxa1.Text = Format(Range("Taxa1"), "0.0000%")
Set rng_Taxa2 = doc.Bookmarks("Taxa2").Range
rng_Taxa2.Text = Format(Range("Taxa2"), "0.0000%")


Set rng_CSC_DTOTAL = doc.Bookmarks("CSC_DTOTAL").Range
rng_CSC_DTOTAL.Text = Format(Range("CSC_DTOTAL"), "0.0000%")

Set rng_CSC_DLINK = doc.Bookmarks("CSC_DLINK").Range
rng_CSC_DLINK.Text = Format(Range("CSC_DLINK"), "0.0000%")

Set rng_CSC_DispTot1 = doc.Bookmarks("CSC_DispTot1").Range
rng_CSC_DispTot1.Text = Range("CSC_DispTot1")

Set rng_CSC_DispTot2 = doc.Bookmarks("CSC_DispTot2").Range
rng_CSC_DispTot2.Text = Range("CSC_DispTot2")

Set rng_CSC_DispTot1p = doc.Bookmarks("CSC_DispTot1p").Range
rng_CSC_DispTot1p.Text = Format(Range("CSC_DispTot1p"), "0.0000%")

Set rng_CSC_DispTot2p = doc.Bookmarks("CSC_DispTot2p").Range
rng_CSC_DispTot2p.Text = Format(Range("CSC_DispTot2p"), "0.0000%")

Set rng_CSC_DispL1 = doc.Bookmarks("CSC_DispL1").Range
rng_CSC_DispL1.Text = Range("CSC_DispL1")

Set rng_CSC_DispL2 = doc.Bookmarks("CSC_DispL2").Range
rng_CSC_DispL2.Text = Range("CSC_DispL2")

Set rng_CSC_DispL1p = doc.Bookmarks("CSC_DispL1p").Range
rng_CSC_DispL1p.Text = Format(Range("CSC_DispL1p"), "0.0000%")

Set rng_CSC_DispL2p = doc.Bookmarks("CSC_DispL2p").Range
rng_CSC_DispL2p.Text = Format(Range("CSC_DispL2p"), "0.0000%")

Set rng_CSC_Prob1 = doc.Bookmarks("CSC_Prob1").Range
rng_CSC_Prob1.Text = Range("CSC_Prob1")

Set rng_CSC_Prob2 = doc.Bookmarks("CSC_Prob2").Range
rng_CSC_Prob2.Text = Range("CSC_Prob2")


Set rng_CC_DTOTAL = doc.Bookmarks("CC_DTOTAL").Range
rng_CC_DTOTAL.Text = Format(Range("CC_DTOTAL"), "0.0000%")

Set rng_CC_DLINK = doc.Bookmarks("CC_DLINK").Range
rng_CC_DLINK.Text = Format(Range("CC_DLINK"), "0.0000%")

Set rng_CC_DispTot1 = doc.Bookmarks("CC_DispTot1").Range
rng_CC_DispTot1.Text = Range("CC_DispTot1")

Set rng_CC_DispTot2 = doc.Bookmarks("CC_DispTot2").Range
rng_CC_DispTot2.Text = Range("CC_DispTot2")

Set rng_CC_DispTot1p = doc.Bookmarks("CC_DispTot1p").Range
rng_CC_DispTot1p.Text = Format(Range("CC_DispTot1p"), "0.0000%")

Set rng_CC_DispTot2p = doc.Bookmarks("CC_DispTot2p").Range
rng_CC_DispTot2p.Text = Format(Range("CC_DispTot2p"), "0.0000%")

Set rng_CC_DispL1 = doc.Bookmarks("CC_DispL1").Range
rng_CC_DispL1.Text = Range("CC_DispL1")

Set rng_CC_DispL2 = doc.Bookmarks("CC_DispL2").Range
rng_CC_DispL2.Text = Range("CC_DispL2")

Set rng_CC_DispL1p = doc.Bookmarks("CC_DispL1p").Range
rng_CC_DispL1p.Text = Format(Range("CC_DispL1p"), "0.0000%")

Set rng_CC_DispL2p = doc.Bookmarks("CC_DispL2p").Range
rng_CC_DispL2p.Text = Format(Range("CC_DispL2p"), "0.0000%")

Set rng_CC_Prob1 = doc.Bookmarks("CC_Prob1").Range
rng_CC_Prob1.Text = Range("CC_Prob1")

Set rng_CC_Prob2 = doc.Bookmarks("CC_Prob2").Range
rng_CC_Prob2.Text = Range("CC_Prob2")


doc.Close
Set doc = Nothing
End Sub
