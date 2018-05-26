Attribute VB_Name = "mdlRibbon"
Option Explicit

Private blnPopupPosition As Boolean   'True:ƒZƒ‹ˆÊ’u  False:’†‰›

Public Sub Ribbon_onLoad_SPEGRA(ribbon As IRibbonUI) ' ƒŠƒ{ƒ“‚Ì‰Šúˆ—
    blnPopupPosition = True
End Sub

Public Sub Ribbon_onAction_SPEGRA(control As IRibbonControl)
    m_SpellAndGrammarChk.selSpellAndGramaCheck
End Sub

