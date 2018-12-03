# -*- coding: mbcs -*-
# Created by makepy.py version 0.4.91
# By python version 2.3.4 (#53, May 25 2004, 21:17:02) [MSC v.1200 32 bit (Intel)]
# From type library 'MSWORD.OLB'
# On Sun Nov 07 07:13:24 2004
"""Microsoft Word 11.0 Object Library"""
makepy_version = '0.4.91'
python_version = 0x20304f0


#CLSID = IID('{00020905-0000-0000-C000-000000000046}')
MajorVersion = 8
MinorVersion = 3
LibraryFlags = 8
LCID = 0x0

#class constants:
wdAlertsAll                   =-1         # from enum WdAlertLevel
wdAlertsMessageBox            =-2         # from enum WdAlertLevel
wdAlertsNone                  =0x0        # from enum WdAlertLevel

wdNumeralArabic               =0x0        # from enum WdArabicNumeral
wdNumeralContext              =0x2        # from enum WdArabicNumeral
wdNumeralHindi                =0x1        # from enum WdArabicNumeral
wdNumeralSystem               =0x3        # from enum WdArabicNumeral

wdIcons                       =0x1        # from enum WdArrangeStyle
wdTiled                       =0x0        # from enum WdArrangeStyle

wdAutoFitContent              =0x1        # from enum WdAutoFitBehavior
wdAutoFitFixed                =0x0        # from enum WdAutoFitBehavior
wdAutoFitWindow               =0x2        # from enum WdAutoFitBehavior

wdAutoClose                   =0x3        # from enum WdAutoMacros
wdAutoExec                    =0x0        # from enum WdAutoMacros
wdAutoExit                    =0x4        # from enum WdAutoMacros
wdAutoNew                     =0x1        # from enum WdAutoMacros
wdAutoOpen                    =0x2        # from enum WdAutoMacros
wdAutoSync                    =0x5        # from enum WdAutoMacros

wdAutoVersionOff              =0x0        # from enum WdAutoVersions
wdAutoVersionOnClose          =0x1        # from enum WdAutoVersions

wdBaselineAlignAuto           =0x4        # from enum WdBaselineAlignment
wdBaselineAlignBaseline       =0x2        # from enum WdBaselineAlignment
wdBaselineAlignCenter         =0x1        # from enum WdBaselineAlignment
wdBaselineAlignFarEast50      =0x3        # from enum WdBaselineAlignment
wdBaselineAlignTop            =0x0        # from enum WdBaselineAlignment

wdSortByLocation              =0x1        # from enum WdBookmarkSortBy
wdSortByName                  =0x0        # from enum WdBookmarkSortBy

wdBorderDistanceFromPageEdge  =0x1        # from enum WdBorderDistanceFrom
wdBorderDistanceFromText      =0x0        # from enum WdBorderDistanceFrom

wdBorderBottom                =-3         # from enum WdBorderType
wdBorderDiagonalDown          =-7         # from enum WdBorderType
wdBorderDiagonalUp            =-8         # from enum WdBorderType
wdBorderHorizontal            =-5         # from enum WdBorderType
wdBorderLeft                  =-2         # from enum WdBorderType
wdBorderRight                 =-4         # from enum WdBorderType
wdBorderTop                   =-1         # from enum WdBorderType
wdBorderVertical              =-6         # from enum WdBorderType

wdColumnBreak                 =0x8        # from enum WdBreakType
wdLineBreak                   =0x6        # from enum WdBreakType
wdLineBreakClearLeft          =0x9        # from enum WdBreakType
wdLineBreakClearRight         =0xa        # from enum WdBreakType
wdPageBreak                   =0x7        # from enum WdBreakType
wdSectionBreakContinuous      =0x3        # from enum WdBreakType
wdSectionBreakEvenPage        =0x4        # from enum WdBreakType
wdSectionBreakNextPage        =0x2        # from enum WdBreakType
wdSectionBreakOddPage         =0x5        # from enum WdBreakType
wdTextWrappingBreak           =0xb        # from enum WdBreakType


wdStyleBlockQuotation         =-85        # from enum WdBuiltinStyle
wdStyleBodyText               =-67        # from enum WdBuiltinStyle
wdStyleBodyText2              =-81        # from enum WdBuiltinStyle
wdStyleBodyText3              =-82        # from enum WdBuiltinStyle
wdStyleBodyTextFirstIndent    =-78        # from enum WdBuiltinStyle
wdStyleBodyTextFirstIndent2   =-79        # from enum WdBuiltinStyle
wdStyleBodyTextIndent         =-68        # from enum WdBuiltinStyle
wdStyleBodyTextIndent2        =-83        # from enum WdBuiltinStyle
wdStyleBodyTextIndent3        =-84        # from enum WdBuiltinStyle
wdStyleCaption                =-35        # from enum WdBuiltinStyle
wdStyleClosing                =-64        # from enum WdBuiltinStyle
wdStyleCommentReference       =-40        # from enum WdBuiltinStyle
wdStyleCommentText            =-31        # from enum WdBuiltinStyle
wdStyleDate                   =-77        # from enum WdBuiltinStyle
wdStyleDefaultParagraphFont   =-66        # from enum WdBuiltinStyle
wdStyleEmphasis               =-89        # from enum WdBuiltinStyle
wdStyleEndnoteReference       =-43        # from enum WdBuiltinStyle
wdStyleEndnoteText            =-44        # from enum WdBuiltinStyle
wdStyleEnvelopeAddress        =-37        # from enum WdBuiltinStyle
wdStyleEnvelopeReturn         =-38        # from enum WdBuiltinStyle
wdStyleFooter                 =-33        # from enum WdBuiltinStyle
wdStyleFootnoteReference      =-39        # from enum WdBuiltinStyle
wdStyleFootnoteText           =-30        # from enum WdBuiltinStyle
wdStyleHeader                 =-32        # from enum WdBuiltinStyle
wdStyleHeading1               =-2         # from enum WdBuiltinStyle
wdStyleHeading2               =-3         # from enum WdBuiltinStyle
wdStyleHeading3               =-4         # from enum WdBuiltinStyle
wdStyleHeading4               =-5         # from enum WdBuiltinStyle
wdStyleHeading5               =-6         # from enum WdBuiltinStyle
wdStyleHeading6               =-7         # from enum WdBuiltinStyle
wdStyleHeading7               =-8         # from enum WdBuiltinStyle
wdStyleHeading8               =-9         # from enum WdBuiltinStyle
wdStyleHeading9               =-10        # from enum WdBuiltinStyle
wdStyleHtmlAcronym            =-96        # from enum WdBuiltinStyle
wdStyleHtmlAddress            =-97        # from enum WdBuiltinStyle
wdStyleHtmlCite               =-98        # from enum WdBuiltinStyle
wdStyleHtmlCode               =-99        # from enum WdBuiltinStyle
wdStyleHtmlDfn                =-100       # from enum WdBuiltinStyle
wdStyleHtmlKbd                =-101       # from enum WdBuiltinStyle
wdStyleHtmlNormal             =-95        # from enum WdBuiltinStyle
wdStyleHtmlPre                =-102       # from enum WdBuiltinStyle
wdStyleHtmlSamp               =-103       # from enum WdBuiltinStyle
wdStyleHtmlTt                 =-104       # from enum WdBuiltinStyle
wdStyleHtmlVar                =-105       # from enum WdBuiltinStyle
wdStyleHyperlink              =-86        # from enum WdBuiltinStyle
wdStyleHyperlinkFollowed      =-87        # from enum WdBuiltinStyle
wdStyleIndex1                 =-11        # from enum WdBuiltinStyle
wdStyleIndex2                 =-12        # from enum WdBuiltinStyle
wdStyleIndex3                 =-13        # from enum WdBuiltinStyle
wdStyleIndex4                 =-14        # from enum WdBuiltinStyle
wdStyleIndex5                 =-15        # from enum WdBuiltinStyle
wdStyleIndex6                 =-16        # from enum WdBuiltinStyle
wdStyleIndex7                 =-17        # from enum WdBuiltinStyle
wdStyleIndex8                 =-18        # from enum WdBuiltinStyle
wdStyleIndex9                 =-19        # from enum WdBuiltinStyle
wdStyleIndexHeading           =-34        # from enum WdBuiltinStyle
wdStyleLineNumber             =-41        # from enum WdBuiltinStyle
wdStyleList                   =-48        # from enum WdBuiltinStyle
wdStyleList2                  =-51        # from enum WdBuiltinStyle
wdStyleList3                  =-52        # from enum WdBuiltinStyle
wdStyleList4                  =-53        # from enum WdBuiltinStyle
wdStyleList5                  =-54        # from enum WdBuiltinStyle
wdStyleListBullet             =-49        # from enum WdBuiltinStyle
wdStyleListBullet2            =-55        # from enum WdBuiltinStyle
wdStyleListBullet3            =-56        # from enum WdBuiltinStyle
wdStyleListBullet4            =-57        # from enum WdBuiltinStyle
wdStyleListBullet5            =-58        # from enum WdBuiltinStyle
wdStyleListContinue           =-69        # from enum WdBuiltinStyle
wdStyleListContinue2          =-70        # from enum WdBuiltinStyle
wdStyleListContinue3          =-71        # from enum WdBuiltinStyle
wdStyleListContinue4          =-72        # from enum WdBuiltinStyle
wdStyleListContinue5          =-73        # from enum WdBuiltinStyle
wdStyleListNumber             =-50        # from enum WdBuiltinStyle
wdStyleListNumber2            =-59        # from enum WdBuiltinStyle
wdStyleListNumber3            =-60        # from enum WdBuiltinStyle
wdStyleListNumber4            =-61        # from enum WdBuiltinStyle
wdStyleListNumber5            =-62        # from enum WdBuiltinStyle
wdStyleMacroText              =-46        # from enum WdBuiltinStyle
wdStyleMessageHeader          =-74        # from enum WdBuiltinStyle
wdStyleNavPane                =-90        # from enum WdBuiltinStyle
wdStyleNormal                 =-1         # from enum WdBuiltinStyle
wdStyleNormalIndent           =-29        # from enum WdBuiltinStyle
wdStyleNormalTable            =-106       # from enum WdBuiltinStyle
wdStyleNoteHeading            =-80        # from enum WdBuiltinStyle
wdStylePageNumber             =-42        # from enum WdBuiltinStyle
wdStylePlainText              =-91        # from enum WdBuiltinStyle
wdStyleSalutation             =-76        # from enum WdBuiltinStyle
wdStyleSignature              =-65        # from enum WdBuiltinStyle
wdStyleStrong                 =-88        # from enum WdBuiltinStyle
wdStyleSubtitle               =-75        # from enum WdBuiltinStyle
wdStyleTOAHeading             =-47        # from enum WdBuiltinStyle
wdStyleTOC1                   =-20        # from enum WdBuiltinStyle
wdStyleTOC2                   =-21        # from enum WdBuiltinStyle
wdStyleTOC3                   =-22        # from enum WdBuiltinStyle
wdStyleTOC4                   =-23        # from enum WdBuiltinStyle
wdStyleTOC5                   =-24        # from enum WdBuiltinStyle
wdStyleTOC6                   =-25        # from enum WdBuiltinStyle
wdStyleTOC7                   =-26        # from enum WdBuiltinStyle
wdStyleTOC8                   =-27        # from enum WdBuiltinStyle
wdStyleTOC9                   =-28        # from enum WdBuiltinStyle
wdStyleTableOfAuthorities     =-45        # from enum WdBuiltinStyle
wdStyleTableOfFigures         =-36        # from enum WdBuiltinStyle
wdStyleTitle                  =-63        # from enum WdBuiltinStyle


wdCaptionEquation             =-3         # from enum WdCaptionLabelID
wdCaptionFigure               =-1         # from enum WdCaptionLabelID
wdCaptionTable                =-2         # from enum WdCaptionLabelID

wdCaptionNumberStyleLowercaseLetter=0x4   # from enum WdCaptionNumberStyle
wdCaptionNumberStyleLowercaseRoman=0x2    # from enum WdCaptionNumberStyle
wdCaptionNumberStyleNumberInCircle=0x12   # from enum WdCaptionNumberStyle
wdCaptionNumberStyleUppercaseLetter=0x3   # from enum WdCaptionNumberStyle
wdCaptionNumberStyleUppercaseRoman=0x1    # from enum WdCaptionNumberStyle
wdCaptionNumberStyleZodiac1   =0x1e       # from enum WdCaptionNumberStyle
wdCaptionNumberStyleZodiac2   =0x1f       # from enum WdCaptionNumberStyle


wdCaptionPositionAbove        =0x0        # from enum WdCaptionPosition
wdCaptionPositionBelow        =0x1        # from enum WdCaptionPosition


wdCellAlignVerticalBottom     =0x3        # from enum WdCellVerticalAlignment
wdCellAlignVerticalCenter     =0x1        # from enum WdCellVerticalAlignment
wdCellAlignVerticalTop        =0x0        # from enum WdCellVerticalAlignment


wdFullWidth                   =0x7        # from enum WdCharacterCase
wdHalfWidth                   =0x6        # from enum WdCharacterCase
wdHiragana                    =0x9        # from enum WdCharacterCase
wdKatakana                    =0x8        # from enum WdCharacterCase
wdLowerCase                   =0x0        # from enum WdCharacterCase
wdNextCase                    =-1         # from enum WdCharacterCase
wdTitleSentence               =0x4        # from enum WdCharacterCase
wdTitleWord                   =0x2        # from enum WdCharacterCase
wdToggleCase                  =0x5        # from enum WdCharacterCase
wdUpperCase                   =0x1        # from enum WdCharacterCase

wdWidthFullWidth              =0x7        # from enum WdCharacterWidth
wdWidthHalfWidth              =0x6        # from enum WdCharacterWidth

# Selection
wdCollapseEnd                 =0x0        # from enum WdCollapseDirection
wdCollapseStart               =0x1        # from enum WdCollapseDirection

# Colors
wdColorAqua                   =0xcccc33   # from enum WdColor
wdColorAutomatic              =-16777216  # from enum WdColor
wdColorBlack                  =0x0        # from enum WdColor
wdColorBlue                   =0xff0000   # from enum WdColor
wdColorBlueGray               =0x996666   # from enum WdColor
wdColorBrightGreen            =0xff00     # from enum WdColor
wdColorBrown                  =0x3399     # from enum WdColor
wdColorDarkBlue               =0x800000   # from enum WdColor
wdColorDarkGreen              =0x3300     # from enum WdColor
wdColorDarkRed                =0x80       # from enum WdColor
wdColorDarkTeal               =0x663300   # from enum WdColor
wdColorDarkYellow             =0x8080     # from enum WdColor
wdColorGold                   =0xccff     # from enum WdColor
wdColorGray05                 =0xf3f3f3   # from enum WdColor
wdColorGray10                 =0xe6e6e6   # from enum WdColor
wdColorGray125                =0xe0e0e0   # from enum WdColor
wdColorGray15                 =0xd9d9d9   # from enum WdColor
wdColorGray20                 =0xcccccc   # from enum WdColor
wdColorGray25                 =0xc0c0c0   # from enum WdColor
wdColorGray30                 =0xb3b3b3   # from enum WdColor
wdColorGray35                 =0xa6a6a6   # from enum WdColor
wdColorGray375                =0xa0a0a0   # from enum WdColor
wdColorGray40                 =0x999999   # from enum WdColor
wdColorGray45                 =0x8c8c8c   # from enum WdColor
wdColorGray50                 =0x808080   # from enum WdColor
wdColorGray55                 =0x737373   # from enum WdColor
wdColorGray60                 =0x666666   # from enum WdColor
wdColorGray625                =0x606060   # from enum WdColor
wdColorGray65                 =0x595959   # from enum WdColor
wdColorGray70                 =0x4c4c4c   # from enum WdColor
wdColorGray75                 =0x404040   # from enum WdColor
wdColorGray80                 =0x333333   # from enum WdColor
wdColorGray85                 =0x262626   # from enum WdColor
wdColorGray875                =0x202020   # from enum WdColor
wdColorGray90                 =0x191919   # from enum WdColor
wdColorGray95                 =0xc0c0c    # from enum WdColor
wdColorGreen                  =0x8000     # from enum WdColor
wdColorIndigo                 =0x993333   # from enum WdColor
wdColorLavender               =0xff99cc   # from enum WdColor
wdColorLightBlue              =0xff6633   # from enum WdColor
wdColorLightGreen             =0xccffcc   # from enum WdColor
wdColorLightOrange            =0x99ff     # from enum WdColor
wdColorLightTurquoise         =0xffffcc   # from enum WdColor
wdColorLightYellow            =0x99ffff   # from enum WdColor
wdColorLime                   =0xcc99     # from enum WdColor
wdColorOliveGreen             =0x3333     # from enum WdColor
wdColorOrange                 =0x66ff     # from enum WdColor
wdColorPaleBlue               =0xffcc99   # from enum WdColor
wdColorPink                   =0xff00ff   # from enum WdColor
wdColorPlum                   =0x663399   # from enum WdColor
wdColorRed                    =0xff       # from enum WdColor
wdColorRose                   =0xcc99ff   # from enum WdColor
wdColorSeaGreen               =0x669933   # from enum WdColor
wdColorSkyBlue                =0xffcc00   # from enum WdColor
wdColorTan                    =0x99ccff   # from enum WdColor
wdColorTeal                   =0x808000   # from enum WdColor
wdColorTurquoise              =0xffff00   # from enum WdColor
wdColorViolet                 =0x800080   # from enum WdColor
wdColorWhite                  =0xffffff   # from enum WdColor
wdColorYellow                 =0xffff     # from enum WdColor

wdAuto                        =0x0        # from enum WdColorIndex
wdBlack                       =0x1        # from enum WdColorIndex
wdBlue                        =0x2        # from enum WdColorIndex
wdBrightGreen                 =0x4        # from enum WdColorIndex
wdByAuthor                    =-1         # from enum WdColorIndex
wdDarkBlue                    =0x9        # from enum WdColorIndex
wdDarkRed                     =0xd        # from enum WdColorIndex
wdDarkYellow                  =0xe        # from enum WdColorIndex
wdGray25                      =0x10       # from enum WdColorIndex
wdGray50                      =0xf        # from enum WdColorIndex
wdGreen                       =0xb        # from enum WdColorIndex
wdNoHighlight                 =0x0        # from enum WdColorIndex
wdPink                        =0x5        # from enum WdColorIndex
wdRed                         =0x6        # from enum WdColorIndex
wdTeal                        =0xa        # from enum WdColorIndex
wdTurquoise                   =0x3        # from enum WdColorIndex
wdViolet                      =0xc        # from enum WdColorIndex
wdWhite                       =0x8        # from enum WdColorIndex
wdYellow                      =0x7        # from enum WdColorIndex

wdCompareTargetCurrent        =0x1        # from enum WdCompareTarget
wdCompareTargetNew            =0x2        # from enum WdCompareTarget
wdCompareTargetSelected       =0x0        # from enum WdCompareTarget


wdAlignTablesRowByRow         =0x27       # from enum WdCompatibility
wdApplyBreakingRules          =0x2e       # from enum WdCompatibility
wdAutospaceLikeWW7            =0x26       # from enum WdCompatibility
wdConvMailMergeEsc            =0x6        # from enum WdCompatibility
wdDontAdjustLineHeightInTable =0x24       # from enum WdCompatibility
wdDontBalanceSingleByteDoubleByteWidth=0x10       # from enum WdCompatibility
wdDontBreakWrappedTables      =0x2b       # from enum WdCompatibility
wdDontSnapTextToGridInTableWithObjects=0x2c       # from enum WdCompatibility
wdDontULTrailSpace            =0xf        # from enum WdCompatibility
wdDontUseAsianBreakRulesInGrid=0x30       # from enum WdCompatibility
wdDontUseHTMLParagraphAutoSpacing=0x23       # from enum WdCompatibility
wdDontWrapTextWithPunctuation =0x2f       # from enum WdCompatibility
wdExactOnTop                  =0x1c       # from enum WdCompatibility
wdExpandShiftReturn           =0xe        # from enum WdCompatibility
wdFootnoteLayoutLikeWW8       =0x22       # from enum WdCompatibility
wdForgetLastTabAlignment      =0x25       # from enum WdCompatibility
wdGrowAutofit                 =0x32       # from enum WdCompatibility
wdLayoutRawTableWidth         =0x28       # from enum WdCompatibility
wdLayoutTableRowsApart        =0x29       # from enum WdCompatibility
wdLeaveBackslashAlone         =0xd        # from enum WdCompatibility
wdLineWrapLikeWord6           =0x20       # from enum WdCompatibility
wdMWSmallCaps                 =0x16       # from enum WdCompatibility
wdNoColumnBalance             =0x5        # from enum WdCompatibility
wdNoExtraLineSpacing          =0x17       # from enum WdCompatibility
wdNoLeading                   =0x14       # from enum WdCompatibility
wdNoSpaceForUL                =0x15       # from enum WdCompatibility
wdNoSpaceRaiseLower           =0x2        # from enum WdCompatibility
wdNoTabHangIndent             =0x1        # from enum WdCompatibility
wdOrigWordTableRules          =0x9        # from enum WdCompatibility
wdPrintBodyTextBeforeHeader   =0x13       # from enum WdCompatibility
wdPrintColBlack               =0x3        # from enum WdCompatibility
wdSelectFieldWithFirstOrLastCharacter=0x2d       # from enum WdCompatibility
wdShapeLayoutLikeWW8          =0x21       # from enum WdCompatibility
wdShowBreaksInFrames          =0xb        # from enum WdCompatibility
wdSpacingInWholePoints        =0x12       # from enum WdCompatibility
wdSubFontBySize               =0x19       # from enum WdCompatibility
wdSuppressBottomSpacing       =0x1d       # from enum WdCompatibility
wdSuppressSpBfAfterPgBrk      =0x7        # from enum WdCompatibility
wdSuppressTopSpacing          =0x8        # from enum WdCompatibility
wdSuppressTopSpacingMac5      =0x11       # from enum WdCompatibility
wdSwapBordersFacingPages      =0xc        # from enum WdCompatibility
wdTransparentMetafiles        =0xa        # from enum WdCompatibility
wdTruncateFontHeight          =0x18       # from enum WdCompatibility
wdUsePrinterMetrics           =0x1a       # from enum WdCompatibility
wdUseWord2002TableStyleRules  =0x31       # from enum WdCompatibility
wdUseWord97LineBreakingRules  =0x2a       # from enum WdCompatibility
wdWPJustification             =0x1f       # from enum WdCompatibility
wdWPSpaceWidth                =0x1e       # from enum WdCompatibility
wdWW6BorderRules              =0x1b       # from enum WdCompatibility
wdWrapTrailSpaces             =0x4        # from enum WdCompatibility


wdEvenColumnBanding           =0x7        # from enum WdConditionCode
wdEvenRowBanding              =0x3        # from enum WdConditionCode
wdFirstColumn                 =0x4        # from enum WdConditionCode
wdFirstRow                    =0x0        # from enum WdConditionCode
wdLastColumn                  =0x5        # from enum WdConditionCode
wdLastRow                     =0x1        # from enum WdConditionCode
wdNECell                      =0x8        # from enum WdConditionCode
wdNWCell                      =0x9        # from enum WdConditionCode
wdOddColumnBanding            =0x6        # from enum WdConditionCode
wdOddRowBanding               =0x2        # from enum WdConditionCode
wdSECell                      =0xa        # from enum WdConditionCode
wdSWCell                      =0xb        # from enum WdConditionCode


wdAutoPosition                =0x0        # from enum WdConstants
wdBackward                    =-1073741823 # from enum WdConstants
wdCreatorCode                 =0x4d535744 # from enum WdConstants
wdFirst                       =0x1        # from enum WdConstants
wdForward                     =0x3fffffff # from enum WdConstants
wdToggle                      =0x98967e   # from enum WdConstants
wdUndefined                   =0x98967f   # from enum WdConstants

wdContinueDisabled            =0x0        # from enum WdContinue
wdContinueList                =0x2        # from enum WdContinue
wdResetList                   =0x1        # from enum WdContinue

wdCursorMovementLogical       =0x0        # from enum WdCursorMovement
wdCursorMovementVisual        =0x1        # from enum WdCursorMovement

wdCursorIBeam                 =0x1        # from enum WdCursorType
wdCursorNormal                =0x2        # from enum WdCursorType
wdCursorNorthwestArrow        =0x3        # from enum WdCursorType
wdCursorWait                  =0x0        # from enum WdCursorType

wdCustomLabelA4               =0x2        # from enum WdCustomLabelPageSize
wdCustomLabelA4LS             =0x3        # from enum WdCustomLabelPageSize
wdCustomLabelA5               =0x4        # from enum WdCustomLabelPageSize
wdCustomLabelA5LS             =0x5        # from enum WdCustomLabelPageSize
wdCustomLabelB4JIS            =0xd        # from enum WdCustomLabelPageSize
wdCustomLabelB5               =0x6        # from enum WdCustomLabelPageSize
wdCustomLabelFanfold          =0x8        # from enum WdCustomLabelPageSize
wdCustomLabelHigaki           =0xb        # from enum WdCustomLabelPageSize
wdCustomLabelHigakiLS         =0xc        # from enum WdCustomLabelPageSize
wdCustomLabelLetter           =0x0        # from enum WdCustomLabelPageSize
wdCustomLabelLetterLS         =0x1        # from enum WdCustomLabelPageSize
wdCustomLabelMini             =0x7        # from enum WdCustomLabelPageSize
wdCustomLabelVertHalfSheet    =0x9        # from enum WdCustomLabelPageSize
wdCustomLabelVertHalfSheetLS  =0xa        # from enum WdCustomLabelPageSize


wdAutoRecoverPath             =0x5        # from enum WdDefaultFilePath
wdBorderArtPath               =0x13       # from enum WdDefaultFilePath
wdCurrentFolderPath           =0xe        # from enum WdDefaultFilePath
wdDocumentsPath               =0x0        # from enum WdDefaultFilePath
wdGraphicsFiltersPath         =0xa        # from enum WdDefaultFilePath
wdPicturesPath                =0x1        # from enum WdDefaultFilePath  ****
wdProgramPath                 =0x9        # from enum WdDefaultFilePath
wdProofingToolsPath           =0xc        # from enum WdDefaultFilePath
wdStartupPath                 =0x8        # from enum WdDefaultFilePath
wdStyleGalleryPath            =0xf        # from enum WdDefaultFilePath
wdTempFilePath                =0xd        # from enum WdDefaultFilePath
wdTextConvertersPath          =0xb        # from enum WdDefaultFilePath
wdToolsPath                   =0x6        # from enum WdDefaultFilePath
wdTutorialPath                =0x7        # from enum WdDefaultFilePath
wdUserOptionsPath             =0x4        # from enum WdDefaultFilePath
wdUserTemplatesPath           =0x2        # from enum WdDefaultFilePath
wdWorkgroupTemplatesPath      =0x3        # from enum WdDefaultFilePath

wdWord10ListBehavior          =0x2        # from enum WdDefaultListBehavior
wdWord8ListBehavior           =0x0        # from enum WdDefaultListBehavior
wdWord9ListBehavior           =0x1        # from enum WdDefaultListBehavior
wdWord8TableBehavior          =0x0        # from enum WdDefaultTableBehavior
wdWord9TableBehavior          =0x1        # from enum WdDefaultTableBehavior

wdDeleteCellsEntireColumn     =0x3        # from enum WdDeleteCells
wdDeleteCellsEntireRow        =0x2        # from enum WdDeleteCells
wdDeleteCellsShiftLeft        =0x0        # from enum WdDeleteCells
wdDeleteCellsShiftUp          =0x1        # from enum WdDeleteCells

wdDeletedTextMarkBold         =0x5        # from enum WdDeletedTextMark
wdDeletedTextMarkCaret        =0x2        # from enum WdDeletedTextMark
wdDeletedTextMarkColorOnly    =0x9        # from enum WdDeletedTextMark
wdDeletedTextMarkDoubleUnderline=0x8        # from enum WdDeletedTextMark
wdDeletedTextMarkHidden       =0x0        # from enum WdDeletedTextMark
wdDeletedTextMarkItalic       =0x6        # from enum WdDeletedTextMark
wdDeletedTextMarkNone         =0x4        # from enum WdDeletedTextMark
wdDeletedTextMarkPound        =0x3        # from enum WdDeletedTextMark
wdDeletedTextMarkStrikeThrough=0x1        # from enum WdDeletedTextMark
wdDeletedTextMarkUnderline    =0x7        # from enum WdDeletedTextMark

wdLeftToRight                 =0x0        # from enum WdDocumentDirection
wdRightToLeft                 =0x1        # from enum WdDocumentDirection

wdDocumentEmail               =0x2        # from enum WdDocumentKind
wdDocumentLetter              =0x1        # from enum WdDocumentKind
wdDocumentNotSpecified        =0x0        # from enum WdDocumentKind

wdDocument                    =0x1        # from enum WdDocumentMedium
wdEmailMessage                =0x0        # from enum WdDocumentMedium
wdWebPage                     =0x2        # from enum WdDocumentMedium

wdTypeDocument                =0x0        # from enum WdDocumentType
wdTypeFrameset                =0x2        # from enum WdDocumentType
wdTypeTemplate                =0x1        # from enum WdDocumentType

wdDocumentViewLtr             =0x1        # from enum WdDocumentViewDirection
wdDocumentViewRtl             =0x0        # from enum WdDocumentViewDirection



wdPropertyTitle	              =0x1        # Title.
wdPropertySubject	          =0x2        # Subject.
wdPropertyAuthor	          =0x3        # Author.
wdPropertyKeywords	          =0x4        # Keywords.
wdPropertyComments	          =0x5        # Comments.
wdPropertyTemplate	          =0x6        # Template name.
wdPropertyLastAuthor	      =0x7        # Last author.
wdPropertyRevision	          =0x8        # Revision number.
wdPropertyAppName	          =0x9        # Name of application.
wdPropertyTimeLastPrinted	  =0xa        # Time last printed.
wdPropertyTimeCreated	      =0xb        # Time created.
wdPropertyTimeLastSaved	      =0xc        # Time last saved.
wdPropertyVBATotalEdit	      =0xd        # Number of edits to VBA project.
wdPropertyPages	              =0xe        # Page count.
wdPropertyWords	              =0xf        # Word count.
wdPropertyCharacters	      =0x10       # Character count.
wdPropertySecurity	          =0x11       # Security setting.
wdPropertyCategory	          =0x12       # Category.
wdPropertyFormat	          =0x13       # Not supported.
wdPropertyManager	          =0x14       # Manager.
wdPropertyCompany	          =0x15       # Company.
wdPropertyBytes	              =0x16       # Byte count.
wdPropertyLines	              =0x17       # Line count.
wdPropertyParas	              =0x18       # Paragraph count.
wdPropertySlides	          =0x19       # Not supported.
wdPropertyNotes	              =0x1a       # Notes.
wdPropertyHiddenSlides	      =0x1b       # Not supported.
wdPropertyMMClips	          =0x1c       # Not supported.
wdPropertyHyperlinkBase	      =0x1d       # Not supported.
wdPropertyCharsWSpaces	      =0x1e       # Character count with spaces.


wdContentControlRichText	        =0x0	# Specifies a rich-text content control.
wdContentControlText	            =0x1	# Specifies a text content control
wdContentControlPicture	            =0x2	# Specifies a picture content control.
wdContentControlComboBox	        =0x3	# Specifies a combo box content control.
wdContentControlDropdownList	    =0x4	# Specifies a drop-down list content control.
wdContentControlBuildingBlockGallery =0x5	# Specifies a building block gallery content control.
wdContentControlDate	            =0x6	# Specifies a date content control.
wdContentControlGroup 	            =0x7	# Specifies a group content control.
wdContentControlCheckbox	        =0x8	# Specifies a checkbox content control.
wdContentControlRepeatingSection	=0x9	# Specifies a repeating section content control.

wdContentControlBoundingBox	        =0x0	# Represents a content control shown as a shaded rectangle or bounding box (with optional title).
wdContentControlTags	            =0x2	# Represents a content control shown as start and end markers.
wdContentControlHidden	            =0x1	# Represents a content control that is not shown.




wdDropMargin                  =0x2        # from enum WdDropPosition
wdDropNone                    =0x0        # from enum WdDropPosition
wdDropNormal                  =0x1        # from enum WdDropPosition

wdBeneathText                 =0x1        # from enum WdFootnoteLocation
wdBottomOfPage                =0x0        # from enum WdFootnoteLocation

wdFrameBottom                 =-999997    # from enum WdFramePosition
wdFrameCenter                 =-999995    # from enum WdFramePosition
wdFrameInside                 =-999994    # from enum WdFramePosition
wdFrameLeft                   =-999998    # from enum WdFramePosition
wdFrameOutside                =-999993    # from enum WdFramePosition
wdFrameRight                  =-999996    # from enum WdFramePosition
wdFrameTop                    =-999999    # from enum WdFramePosition

wdFrameAtLeast                =0x1        # from enum WdFrameSizeRule
wdFrameAuto                   =0x0        # from enum WdFrameSizeRule
wdFrameExact                  =0x2        # from enum WdFrameSizeRule

wdFramesetNewFrameAbove       =0x0        # from enum WdFramesetNewFrameLocation
wdFramesetNewFrameBelow       =0x1        # from enum WdFramesetNewFrameLocation
wdFramesetNewFrameLeft        =0x3        # from enum WdFramesetNewFrameLocation
wdFramesetNewFrameRight       =0x2        # from enum WdFramesetNewFrameLocation

wdFramesetSizeTypeFixed       =0x1        # from enum WdFramesetSizeType
wdFramesetSizeTypePercent     =0x0        # from enum WdFramesetSizeType
wdFramesetSizeTypeRelative    =0x2        # from enum WdFramesetSizeType

wdFramesetTypeFrame           =0x1        # from enum WdFramesetType
wdFramesetTypeFrameset        =0x0        # from enum WdFramesetType

# GOTO directions
wdGoToAbsolute                =0x1        # from enum WdGoToDirection
wdGoToFirst                   =0x1        # from enum WdGoToDirection
wdGoToLast                    =-1         # from enum WdGoToDirection
wdGoToNext                    =0x2        # from enum WdGoToDirection
wdGoToPrevious                =0x3        # from enum WdGoToDirection
wdGoToRelative                =0x2        # from enum WdGoToDirection

wdGoToBookmark                =-1         # from enum WdGoToItem
wdGoToComment                 =0x6        # from enum WdGoToItem
wdGoToEndnote                 =0x5        # from enum WdGoToItem
wdGoToEquation                =0xa        # from enum WdGoToItem
wdGoToField                   =0x7        # from enum WdGoToItem
wdGoToFootnote                =0x4        # from enum WdGoToItem
wdGoToGrammaticalError        =0xe        # from enum WdGoToItem
wdGoToGraphic                 =0x8        # from enum WdGoToItem
wdGoToHeading                 =0xb        # from enum WdGoToItem
wdGoToLine                    =0x3        # from enum WdGoToItem
wdGoToObject                  =0x9        # from enum WdGoToItem
wdGoToPage                    =0x1        # from enum WdGoToItem
wdGoToPercent                 =0xc        # from enum WdGoToItem
wdGoToProofreadingError       =0xf        # from enum WdGoToItem
wdGoToSection                 =0x0        # from enum WdGoToItem
wdGoToSpellingError           =0xd        # from enum WdGoToItem
wdGoToTable                   =0x2        # from enum WdGoToItem

# page setup
wdGutterPosLeft               =0x0        # from enum WdGutterStyle
wdGutterPosRight              =0x2        # from enum WdGutterStyle
wdGutterPosTop                =0x1        # from enum WdGutterStyle
wdGutterStyleBidi             =0x2        # from enum WdGutterStyleOld
wdGutterStyleLatin            =-10        # from enum WdGutterStyleOld

wdHeaderFooterEvenPages       =0x3        # from enum WdHeaderFooterIndex
wdHeaderFooterFirstPage       =0x2        # from enum WdHeaderFooterIndex
wdHeaderFooterPrimary         =0x1        # from enum WdHeaderFooterIndex

wdHeadingSeparatorBlankLine   =0x1        # from enum WdHeadingSeparator
wdHeadingSeparatorLetter      =0x2        # from enum WdHeadingSeparator
wdHeadingSeparatorLetterFull  =0x4        # from enum WdHeadingSeparator
wdHeadingSeparatorLetterLow   =0x3        # from enum WdHeadingSeparator
wdHeadingSeparatorNone        =0x0        # from enum WdHeadingSeparator

wdHorizontalInVerticalFitInLine=0x1       # from enum WdHorizontalInVerticalType
wdHorizontalInVerticalNone    =0x0        # from enum WdHorizontalInVerticalType
wdHorizontalInVerticalResizeLine=0x2      # from enum WdHorizontalInVerticalType

wdHorizontalLineAlignCenter   =0x1        # from enum WdHorizontalLineAlignment
wdHorizontalLineAlignLeft     =0x0        # from enum WdHorizontalLineAlignment
wdHorizontalLineAlignRight    =0x2        # from enum WdHorizontalLineAlignment
wdHorizontalLineFixedWidth    =-2         # from enum WdHorizontalLineWidthType
wdHorizontalLinePercentWidth  =-1         # from enum WdHorizontalLineWidthType

wdIndexFilterFull             =0x6        # from enum WdIndexFilter
wdIndexFilterLow              =0x4        # from enum WdIndexFilter
wdIndexFilterMedium           =0x5        # from enum WdIndexFilter
wdIndexFilterNone             =0x0        # from enum WdIndexFilter
wdIndexBulleted               =0x4        # from enum WdIndexFormat
wdIndexClassic                =0x1        # from enum WdIndexFormat
wdIndexFancy                  =0x2        # from enum WdIndexFormat
wdIndexFormal                 =0x5        # from enum WdIndexFormat
wdIndexModern                 =0x3        # from enum WdIndexFormat
wdIndexSimple                 =0x6        # from enum WdIndexFormat
wdIndexTemplate               =0x0        # from enum WdIndexFormat
wdIndexSortByStroke           =0x0        # from enum WdIndexSortBy
wdIndexSortBySyllable         =0x1        # from enum WdIndexSortBy

wdIndexIndent                 =0x0        # from enum WdIndexType
wdIndexRunin                  =0x1        # from enum WdIndexType

wdActiveEndAdjustedPageNumber =0x1        # from enum WdInformation
wdActiveEndPageNumber         =0x3        # from enum WdInformation
wdActiveEndSectionNumber      =0x2        # from enum WdInformation
wdAtEndOfRowMarker            =0x1f       # from enum WdInformation
wdCapsLock                    =0x15       # from enum WdInformation
wdEndOfRangeColumnNumber      =0x11       # from enum WdInformation
wdEndOfRangeRowNumber         =0xe        # from enum WdInformation
wdFirstCharacterColumnNumber  =0x9        # from enum WdInformation
wdFirstCharacterLineNumber    =0xa        # from enum WdInformation
wdFrameIsSelected             =0xb        # from enum WdInformation
wdHeaderFooterType            =0x21       # from enum WdInformation
wdHorizontalPositionRelativeToPage=0x5        # from enum WdInformation
wdHorizontalPositionRelativeToTextBoundary=0x7        # from enum WdInformation
wdInClipboard                 =0x26       # from enum WdInformation
wdInCommentPane               =0x1a       # from enum WdInformation
wdInEndnote                   =0x24       # from enum WdInformation
wdInFootnote                  =0x23       # from enum WdInformation
wdInFootnoteEndnotePane       =0x19       # from enum WdInformation
wdInHeaderFooter              =0x1c       # from enum WdInformation
wdInMasterDocument            =0x22       # from enum WdInformation
wdInWordMail                  =0x25       # from enum WdInformation
wdMaximumNumberOfColumns      =0x12       # from enum WdInformation
wdMaximumNumberOfRows         =0xf        # from enum WdInformation
wdNumLock                     =0x16       # from enum WdInformation
wdNumberOfPagesInDocument     =0x4        # from enum WdInformation
wdOverType                    =0x17       # from enum WdInformation
wdReferenceOfType             =0x20       # from enum WdInformation
wdRevisionMarking             =0x18       # from enum WdInformation
wdSelectionMode               =0x14       # from enum WdInformation
wdStartOfRangeColumnNumber    =0x10       # from enum WdInformation
wdStartOfRangeRowNumber       =0xd        # from enum WdInformation
wdVerticalPositionRelativeToPage=0x6        # from enum WdInformation
wdVerticalPositionRelativeToTextBoundary=0x8        # from enum WdInformation
wdWithInTable                 =0xc        # from enum WdInformation
wdZoomPercentage              =0x13       # from enum WdInformation

wdInlineShapeEmbeddedOLEObject=0x1        # from enum WdInlineShapeType
wdInlineShapeHorizontalLine   =0x6        # from enum WdInlineShapeType
wdInlineShapeLinkedOLEObject  =0x2        # from enum WdInlineShapeType
wdInlineShapeLinkedPicture    =0x4        # from enum WdInlineShapeType
wdInlineShapeLinkedPictureHorizontalLine=0x8        # from enum WdInlineShapeType
wdInlineShapeOLEControlObject =0x5        # from enum WdInlineShapeType
wdInlineShapeOWSAnchor        =0xb        # from enum WdInlineShapeType
wdInlineShapePicture          =0x3        # from enum WdInlineShapeType
wdInlineShapePictureBullet    =0x9        # from enum WdInlineShapeType
wdInlineShapePictureHorizontalLine=0x7        # from enum WdInlineShapeType
wdInlineShapeScriptAnchor     =0xa        # from enum WdInlineShapeType

wdInsertCellsEntireColumn     =0x3        # from enum WdInsertCells
wdInsertCellsEntireRow        =0x2        # from enum WdInsertCells
wdInsertCellsShiftDown        =0x1        # from enum WdInsertCells
wdInsertCellsShiftRight       =0x0        # from enum WdInsertCells

wdInsertedTextMarkBold        =0x1        # from enum WdInsertedTextMark
wdInsertedTextMarkColorOnly   =0x5        # from enum WdInsertedTextMark
wdInsertedTextMarkDoubleUnderline=0x4        # from enum WdInsertedTextMark
wdInsertedTextMarkItalic      =0x2        # from enum WdInsertedTextMark
wdInsertedTextMarkNone        =0x0        # from enum WdInsertedTextMark
wdInsertedTextMarkStrikeThrough=0x6        # from enum WdInsertedTextMark
wdInsertedTextMarkUnderline   =0x3        # from enum WdInsertedTextMark

wdKey0                        =0x30       # from enum WdKey
wdKey1                        =0x31       # from enum WdKey
wdKey2                        =0x32       # from enum WdKey
wdKey3                        =0x33       # from enum WdKey
wdKey4                        =0x34       # from enum WdKey
wdKey5                        =0x35       # from enum WdKey
wdKey6                        =0x36       # from enum WdKey
wdKey7                        =0x37       # from enum WdKey
wdKey8                        =0x38       # from enum WdKey
wdKey9                        =0x39       # from enum WdKey
wdKeyA                        =0x41       # from enum WdKey
wdKeyAlt                      =0x400      # from enum WdKey
wdKeyB                        =0x42       # from enum WdKey
wdKeyBackSingleQuote          =0xc0       # from enum WdKey
wdKeyBackSlash                =0xdc       # from enum WdKey
wdKeyBackspace                =0x8        # from enum WdKey
wdKeyC                        =0x43       # from enum WdKey
wdKeyCloseSquareBrace         =0xdd       # from enum WdKey
wdKeyComma                    =0xbc       # from enum WdKey
wdKeyCommand                  =0x200      # from enum WdKey
wdKeyControl                  =0x200      # from enum WdKey
wdKeyD                        =0x44       # from enum WdKey
wdKeyDelete                   =0x2e       # from enum WdKey
wdKeyE                        =0x45       # from enum WdKey
wdKeyEnd                      =0x23       # from enum WdKey
wdKeyEquals                   =0xbb       # from enum WdKey
wdKeyEsc                      =0x1b       # from enum WdKey
wdKeyF                        =0x46       # from enum WdKey
wdKeyF1                       =0x70       # from enum WdKey
wdKeyF10                      =0x79       # from enum WdKey
wdKeyF11                      =0x7a       # from enum WdKey
wdKeyF12                      =0x7b       # from enum WdKey
wdKeyF13                      =0x7c       # from enum WdKey
wdKeyF14                      =0x7d       # from enum WdKey
wdKeyF15                      =0x7e       # from enum WdKey
wdKeyF16                      =0x7f       # from enum WdKey
wdKeyF2                       =0x71       # from enum WdKey
wdKeyF3                       =0x72       # from enum WdKey
wdKeyF4                       =0x73       # from enum WdKey
wdKeyF5                       =0x74       # from enum WdKey
wdKeyF6                       =0x75       # from enum WdKey
wdKeyF7                       =0x76       # from enum WdKey
wdKeyF8                       =0x77       # from enum WdKey
wdKeyF9                       =0x78       # from enum WdKey
wdKeyG                        =0x47       # from enum WdKey
wdKeyH                        =0x48       # from enum WdKey
wdKeyHome                     =0x24       # from enum WdKey
wdKeyHyphen                   =0xbd       # from enum WdKey
wdKeyI                        =0x49       # from enum WdKey
wdKeyInsert                   =0x2d       # from enum WdKey
wdKeyJ                        =0x4a       # from enum WdKey
wdKeyK                        =0x4b       # from enum WdKey
wdKeyL                        =0x4c       # from enum WdKey
wdKeyM                        =0x4d       # from enum WdKey
wdKeyN                        =0x4e       # from enum WdKey
wdKeyNumeric0                 =0x60       # from enum WdKey
wdKeyNumeric1                 =0x61       # from enum WdKey
wdKeyNumeric2                 =0x62       # from enum WdKey
wdKeyNumeric3                 =0x63       # from enum WdKey
wdKeyNumeric4                 =0x64       # from enum WdKey
wdKeyNumeric5                 =0x65       # from enum WdKey
wdKeyNumeric5Special          =0xc        # from enum WdKey
wdKeyNumeric6                 =0x66       # from enum WdKey
wdKeyNumeric7                 =0x67       # from enum WdKey
wdKeyNumeric8                 =0x68       # from enum WdKey
wdKeyNumeric9                 =0x69       # from enum WdKey
wdKeyNumericAdd               =0x6b       # from enum WdKey
wdKeyNumericDecimal           =0x6e       # from enum WdKey
wdKeyNumericDivide            =0x6f       # from enum WdKey
wdKeyNumericMultiply          =0x6a       # from enum WdKey
wdKeyNumericSubtract          =0x6d       # from enum WdKey
wdKeyO                        =0x4f       # from enum WdKey
wdKeyOpenSquareBrace          =0xdb       # from enum WdKey
wdKeyOption                   =0x400      # from enum WdKey
wdKeyP                        =0x50       # from enum WdKey
wdKeyPageDown                 =0x22       # from enum WdKey
wdKeyPageUp                   =0x21       # from enum WdKey
wdKeyPause                    =0x13       # from enum WdKey
wdKeyPeriod                   =0xbe       # from enum WdKey
wdKeyQ                        =0x51       # from enum WdKey
wdKeyR                        =0x52       # from enum WdKey
wdKeyReturn                   =0xd        # from enum WdKey
wdKeyS                        =0x53       # from enum WdKey
wdKeyScrollLock               =0x91       # from enum WdKey
wdKeySemiColon                =0xba       # from enum WdKey
wdKeyShift                    =0x100      # from enum WdKey
wdKeySingleQuote              =0xde       # from enum WdKey
wdKeySlash                    =0xbf       # from enum WdKey
wdKeySpacebar                 =0x20       # from enum WdKey
wdKeyT                        =0x54       # from enum WdKey
wdKeyTab                      =0x9        # from enum WdKey
wdKeyU                        =0x55       # from enum WdKey
wdKeyV                        =0x56       # from enum WdKey
wdKeyW                        =0x57       # from enum WdKey
wdKeyX                        =0x58       # from enum WdKey
wdKeyY                        =0x59       # from enum WdKey
wdKeyZ                        =0x5a       # from enum WdKey
wdNoKey                       =0xff       # from enum WdKey

wdKeyCategoryAutoText         =0x4        # from enum WdKeyCategory
wdKeyCategoryCommand          =0x1        # from enum WdKeyCategory
wdKeyCategoryDisable          =0x0        # from enum WdKeyCategory
wdKeyCategoryFont             =0x3        # from enum WdKeyCategory
wdKeyCategoryMacro            =0x2        # from enum WdKeyCategory
wdKeyCategoryNil              =-1         # from enum WdKeyCategory
wdKeyCategoryPrefix           =0x7        # from enum WdKeyCategory
wdKeyCategoryStyle            =0x5        # from enum WdKeyCategory
wdKeyCategorySymbol           =0x6        # from enum WdKeyCategory


wdLayoutModeDefault           =0x0        # from enum WdLayoutMode
wdLayoutModeGenko             =0x3        # from enum WdLayoutMode
wdLayoutModeGrid              =0x1        # from enum WdLayoutMode
wdLayoutModeLineGrid          =0x2        # from enum WdLayoutMode

wdFullBlock                   =0x0        # from enum WdLetterStyle
wdModifiedBlock               =0x1        # from enum WdLetterStyle
wdSemiBlock                   =0x2        # from enum WdLetterStyle

wdLetterBottom                =0x1        # from enum WdLetterheadLocation
wdLetterLeft                  =0x2        # from enum WdLetterheadLocation
wdLetterRight                 =0x3        # from enum WdLetterheadLocation
wdLetterTop                   =0x0        # from enum WdLetterheadLocation

wdCRLF                        =0x0        # from enum WdLineEndingType
wdCROnly                      =0x1        # from enum WdLineEndingType
wdLFCR                        =0x3        # from enum WdLineEndingType
wdLFOnly                      =0x2        # from enum WdLineEndingType
wdLSPS                        =0x4        # from enum WdLineEndingType

wdLineSpace1pt5               =0x1        # from enum WdLineSpacing
wdLineSpaceAtLeast            =0x3        # from enum WdLineSpacing
wdLineSpaceDouble             =0x2        # from enum WdLineSpacing
wdLineSpaceExactly            =0x4        # from enum WdLineSpacing
wdLineSpaceMultiple           =0x5        # from enum WdLineSpacing
wdLineSpaceSingle             =0x0        # from enum WdLineSpacing

# Line Styles
wdLineStyleDashDot            =0x5        # from enum WdLineStyle
wdLineStyleDashDotDot         =0x6        # from enum WdLineStyle
wdLineStyleDashDotStroked     =0x14       # from enum WdLineStyle
wdLineStyleDashLargeGap       =0x4        # from enum WdLineStyle
wdLineStyleDashSmallGap       =0x3        # from enum WdLineStyle
wdLineStyleDot                =0x2        # from enum WdLineStyle
wdLineStyleDouble             =0x7        # from enum WdLineStyle
wdLineStyleDoubleWavy         =0x13       # from enum WdLineStyle
wdLineStyleEmboss3D           =0x15       # from enum WdLineStyle
wdLineStyleEngrave3D          =0x16       # from enum WdLineStyle
wdLineStyleInset              =0x18       # from enum WdLineStyle
wdLineStyleNone               =0x0        # from enum WdLineStyle
wdLineStyleOutset             =0x17       # from enum WdLineStyle
wdLineStyleSingle             =0x1        # from enum WdLineStyle
wdLineStyleSingleWavy         =0x12       # from enum WdLineStyle
wdLineStyleThickThinLargeGap  =0x10       # from enum WdLineStyle
wdLineStyleThickThinMedGap    =0xd        # from enum WdLineStyle
wdLineStyleThickThinSmallGap  =0xa        # from enum WdLineStyle
wdLineStyleThinThickLargeGap  =0xf        # from enum WdLineStyle
wdLineStyleThinThickMedGap    =0xc        # from enum WdLineStyle
wdLineStyleThinThickSmallGap  =0x9        # from enum WdLineStyle
wdLineStyleThinThickThinLargeGap=0x11       # from enum WdLineStyle
wdLineStyleThinThickThinMedGap=0xe        # from enum WdLineStyle
wdLineStyleThinThickThinSmallGap=0xb        # from enum WdLineStyle
wdLineStyleTriple             =0x8        # from enum WdLineStyle

wdTableRow                    =0x1        # from enum WdLineType
wdTextLine                    =0x0        # from enum WdLineType

wdLineWidth025pt              =0x2        # from enum WdLineWidth
wdLineWidth050pt              =0x4        # from enum WdLineWidth
wdLineWidth075pt              =0x6        # from enum WdLineWidth
wdLineWidth100pt              =0x8        # from enum WdLineWidth
wdLineWidth150pt              =0xc        # from enum WdLineWidth
wdLineWidth225pt              =0x12       # from enum WdLineWidth
wdLineWidth300pt              =0x18       # from enum WdLineWidth
wdLineWidth450pt              =0x24       # from enum WdLineWidth
wdLineWidth600pt              =0x30       # from enum WdLineWidth

wdLinkTypeDDE                 =0x6        # from enum WdLinkType
wdLinkTypeDDEAuto             =0x7        # from enum WdLinkType
wdLinkTypeImport              =0x5        # from enum WdLinkType
wdLinkTypeInclude             =0x4        # from enum WdLinkType
wdLinkTypeOLE                 =0x0        # from enum WdLinkType
wdLinkTypePicture             =0x1        # from enum WdLinkType
wdLinkTypeReference           =0x3        # from enum WdLinkType
wdLinkTypeText                =0x2        # from enum WdLinkType

wdListApplyToSelection        =0x2        # from enum WdListApplyTo
wdListApplyToThisPointForward =0x1        # from enum WdListApplyTo
wdListApplyToWholeList        =0x0        # from enum WdListApplyTo

wdListLevelAlignCenter        =0x1        # from enum WdListLevelAlignment
wdListLevelAlignLeft          =0x0        # from enum WdListLevelAlignment
wdListLevelAlignRight         =0x2        # from enum WdListLevelAlignment

wdListNumberStyleBullet       =0x17       # from enum WdListNumberStyle
wdListNumberStyleCardinalText =0x6        # from enum WdListNumberStyle
wdListNumberStyleLowercaseLetter=0x4      # from enum WdListNumberStyle
wdListNumberStyleLowercaseRoman=0x2       # from enum WdListNumberStyle
wdListNumberStyleLowercaseRussian=0x3a    # from enum WdListNumberStyle
wdListNumberStyleNone         =0xff       # from enum WdListNumberStyle
wdListNumberStyleNumberInCircle=0x12      # from enum WdListNumberStyle
wdListNumberStyleOrdinal      =0x5        # from enum WdListNumberStyle
wdListNumberStyleOrdinalText  =0x7        # from enum WdListNumberStyle
wdListNumberStylePictureBullet=0xf9       # from enum WdListNumberStyle
wdListNumberStyleUppercaseLetter=0x3      # from enum WdListNumberStyle

wdListBullet                  =0x2        # from enum WdListType
wdListListNumOnly             =0x1        # from enum WdListType
wdListMixedNumbering          =0x5        # from enum WdListType
wdListNoNumbering             =0x0        # from enum WdListType
wdListOutlineNumbering        =0x4        # from enum WdListType
wdListPictureBullet           =0x6        # from enum WdListType
wdListSimpleNumbering         =0x3        # from enum WdListType



wdNewBlankDocument            =0x0        # from enum WdNewDocumentType
wdNewEmailMessage             =0x2        # from enum WdNewDocumentType
wdNewFrameset                 =0x3        # from enum WdNewDocumentType
wdNewWebPage                  =0x1        # from enum WdNewDocumentType
wdNewXMLDocument              =0x4        # from enum WdNewDocumentType

wdNoteNumberStyleLowercaseLetter=0x4      # from enum WdNoteNumberStyle
wdNoteNumberStyleLowercaseRoman=0x2       # from enum WdNoteNumberStyle
wdNoteNumberStyleUppercaseLetter=0x3      # from enum WdNoteNumberStyle
wdNoteNumberStyleUppercaseRoman=0x1       # from enum WdNoteNumberStyle

wdNumberAllNumbers            =0x3        # from enum WdNumberType
wdNumberListNum               =0x2        # from enum WdNumberType
wdNumberParagraph             =0x1        # from enum WdNumberType
wdRestartContinuous           =0x0        # from enum WdNumberingRule
wdRestartPage                 =0x2        # from enum WdNumberingRule
wdRestartSection              =0x1        # from enum WdNumberingRule

wdFloatOverText               =0x1        # from enum WdOLEPlacement
wdInLine                      =0x0        # from enum WdOLEPlacement

wdOLEControl                  =0x2        # from enum WdOLEType
wdOLEEmbed                    =0x1        # from enum WdOLEType
wdOLELink                     =0x0        # from enum WdOLEType
wdOLEVerbDiscardUndoState     =-6         # from enum WdOLEVerb
wdOLEVerbHide                 =-3         # from enum WdOLEVerb
wdOLEVerbInPlaceActivate      =-5         # from enum WdOLEVerb
wdOLEVerbOpen                 =-2         # from enum WdOLEVerb
wdOLEVerbPrimary              =0x0        # from enum WdOLEVerb
wdOLEVerbShow                 =-1         # from enum WdOLEVerb
wdOLEVerbUIActivate           =-4         # from enum WdOLEVerb

wdOpenFormatAllWord           =0x6        # from enum WdOpenFormat
wdOpenFormatAuto              =0x0        # from enum WdOpenFormat
wdOpenFormatDocument          =0x1        # from enum WdOpenFormat
wdOpenFormatEncodedText       =0x5        # from enum WdOpenFormat
wdOpenFormatRTF               =0x3        # from enum WdOpenFormat
wdOpenFormatTemplate          =0x2        # from enum WdOpenFormat
wdOpenFormatText              =0x4        # from enum WdOpenFormat
wdOpenFormatUnicodeText       =0x5        # from enum WdOpenFormat
wdOpenFormatWebPages          =0x7        # from enum WdOpenFormat
wdOpenFormatXML               =0x8        # from enum WdOpenFormat

wdOrganizerObjectAutoText     =0x1        # from enum WdOrganizerObject
wdOrganizerObjectCommandBars  =0x2        # from enum WdOrganizerObject
wdOrganizerObjectProjectItems =0x3        # from enum WdOrganizerObject
wdOrganizerObjectStyles       =0x0        # from enum WdOrganizerObject

wdOrientLandscape             =0x1        # from enum WdOrientation
wdOrientPortrait              =0x0        # from enum WdOrientation

wdOriginalDocumentFormat      =0x1        # from enum WdOriginalFormat
wdPromptUser                  =0x2        # from enum WdOriginalFormat
wdWordDocument                =0x0        # from enum WdOriginalFormat

wdOutlineLevel1               =0x1        # from enum WdOutlineLevel
wdOutlineLevel2               =0x2        # from enum WdOutlineLevel
wdOutlineLevel3               =0x3        # from enum WdOutlineLevel
wdOutlineLevel4               =0x4        # from enum WdOutlineLevel
wdOutlineLevel5               =0x5        # from enum WdOutlineLevel
wdOutlineLevel6               =0x6        # from enum WdOutlineLevel
wdOutlineLevel7               =0x7        # from enum WdOutlineLevel
wdOutlineLevel8               =0x8        # from enum WdOutlineLevel
wdOutlineLevel9               =0x9        # from enum WdOutlineLevel
wdOutlineLevelBodyText        =0xa        # from enum WdOutlineLevel

wdPageFitBestFit              =0x2        # from enum WdPageFit
wdPageFitFullPage             =0x1        # from enum WdPageFit
wdPageFitNone                 =0x0        # from enum WdPageFit
wdPageFitTextFit              =0x3        # from enum WdPageFit

wdAlignPageNumberCenter       =0x1        # from enum WdPageNumberAlignment
wdAlignPageNumberInside       =0x3        # from enum WdPageNumberAlignment
wdAlignPageNumberLeft         =0x0        # from enum WdPageNumberAlignment
wdAlignPageNumberOutside      =0x4        # from enum WdPageNumberAlignment
wdAlignPageNumberRight        =0x2        # from enum WdPageNumberAlignment

wdPageNumberStyleLowercaseLetter=0x4        # from enum WdPageNumberStyle
wdPageNumberStyleUppercaseLetter=0x3        # from enum WdPageNumberStyle
wdPageNumberStyleUppercaseRoman=0x1        # from enum WdPageNumberStyle

wdPaper10x14                  =0x0        # from enum WdPaperSize
wdPaper11x17                  =0x1        # from enum WdPaperSize
wdPaperA3                     =0x6        # from enum WdPaperSize
wdPaperA4                     =0x7        # from enum WdPaperSize
wdPaperA4Small                =0x8        # from enum WdPaperSize
wdPaperA5                     =0x9        # from enum WdPaperSize
wdPaperB4                     =0xa        # from enum WdPaperSize
wdPaperB5                     =0xb        # from enum WdPaperSize
wdPaperCSheet                 =0xc        # from enum WdPaperSize
wdPaperCustom                 =0x29       # from enum WdPaperSize
wdPaperDSheet                 =0xd        # from enum WdPaperSize
wdPaperESheet                 =0xe        # from enum WdPaperSize
wdPaperEnvelope10             =0x19       # from enum WdPaperSize
wdPaperEnvelope11             =0x1a       # from enum WdPaperSize
wdPaperEnvelope12             =0x1b       # from enum WdPaperSize
wdPaperEnvelope14             =0x1c       # from enum WdPaperSize
wdPaperEnvelope9              =0x18       # from enum WdPaperSize
wdPaperEnvelopeB4             =0x1d       # from enum WdPaperSize
wdPaperEnvelopeB5             =0x1e       # from enum WdPaperSize
wdPaperEnvelopeB6             =0x1f       # from enum WdPaperSize
wdPaperEnvelopeC3             =0x20       # from enum WdPaperSize
wdPaperEnvelopeC4             =0x21       # from enum WdPaperSize
wdPaperEnvelopeC5             =0x22       # from enum WdPaperSize
wdPaperEnvelopeC6             =0x23       # from enum WdPaperSize
wdPaperEnvelopeC65            =0x24       # from enum WdPaperSize
wdPaperEnvelopeDL             =0x25       # from enum WdPaperSize
wdPaperEnvelopeItaly          =0x26       # from enum WdPaperSize
wdPaperEnvelopeMonarch        =0x27       # from enum WdPaperSize
wdPaperEnvelopePersonal       =0x28       # from enum WdPaperSize
wdPaperExecutive              =0x5        # from enum WdPaperSize
wdPaperFanfoldLegalGerman     =0xf        # from enum WdPaperSize
wdPaperFanfoldStdGerman       =0x10       # from enum WdPaperSize
wdPaperFanfoldUS              =0x11       # from enum WdPaperSize
wdPaperFolio                  =0x12       # from enum WdPaperSize
wdPaperLedger                 =0x13       # from enum WdPaperSize
wdPaperLegal                  =0x4        # from enum WdPaperSize
wdPaperLetter                 =0x2        # from enum WdPaperSize
wdPaperLetterSmall            =0x3        # from enum WdPaperSize
wdPaperNote                   =0x14       # from enum WdPaperSize
wdPaperQuarto                 =0x15       # from enum WdPaperSize
wdPaperStatement              =0x16       # from enum WdPaperSize
wdPaperTabloid                =0x17       # from enum WdPaperSize
wdPrinterAutomaticSheetFeed   =0x7        # from enum WdPaperTray
wdPrinterDefaultBin           =0x0        # from enum WdPaperTray
wdPrinterEnvelopeFeed         =0x5        # from enum WdPaperTray
wdPrinterFormSource           =0xf        # from enum WdPaperTray
wdPrinterLargeCapacityBin     =0xb        # from enum WdPaperTray
wdPrinterLargeFormatBin       =0xa        # from enum WdPaperTray
wdPrinterLowerBin             =0x2        # from enum WdPaperTray
wdPrinterManualEnvelopeFeed   =0x6        # from enum WdPaperTray
wdPrinterManualFeed           =0x4        # from enum WdPaperTray
wdPrinterMiddleBin            =0x3        # from enum WdPaperTray
wdPrinterOnlyBin              =0x1        # from enum WdPaperTray
wdPrinterPaperCassette        =0xe        # from enum WdPaperTray
wdPrinterSmallFormatBin       =0x9        # from enum WdPaperTray
wdPrinterTractorFeed          =0x8        # from enum WdPaperTray
wdPrinterUpperBin             =0x1        # from enum WdPaperTray

wdAlignParagraphCenter        =0x1        # from enum WdParagraphAlignment
wdAlignParagraphDistribute    =0x4        # from enum WdParagraphAlignment
wdAlignParagraphJustify       =0x3        # from enum WdParagraphAlignment
wdAlignParagraphJustifyHi     =0x7        # from enum WdParagraphAlignment
wdAlignParagraphJustifyLow    =0x8        # from enum WdParagraphAlignment
wdAlignParagraphJustifyMed    =0x5        # from enum WdParagraphAlignment
wdAlignParagraphLeft          =0x0        # from enum WdParagraphAlignment
wdAlignParagraphRight         =0x2        # from enum WdParagraphAlignment
wdAlignParagraphThaiJustify   =0x9        # from enum WdParagraphAlignment

wdPasteBitmap                 =0x4        # from enum WdPasteDataType
wdPasteDeviceIndependentBitmap=0x5        # from enum WdPasteDataType
wdPasteEnhancedMetafile       =0x9        # from enum WdPasteDataType
wdPasteHTML                   =0xa        # from enum WdPasteDataType
wdPasteHyperlink              =0x7        # from enum WdPasteDataType
wdPasteMetafilePicture        =0x3        # from enum WdPasteDataType
wdPasteOLEObject              =0x0        # from enum WdPasteDataType
wdPasteRTF                    =0x1        # from enum WdPasteDataType
wdPasteShape                  =0x8        # from enum WdPasteDataType
wdPasteText                   =0x2        # from enum WdPasteDataType

wdLinkDataInDoc               =0x1        # from enum WdPictureLinkType
wdLinkDataOnDisk              =0x2        # from enum WdPictureLinkType
wdLinkNone                    =0x0        # from enum WdPictureLinkType

wdPreferredWidthAuto          =0x1        # from enum WdPreferredWidthType
wdPreferredWidthPercent       =0x2        # from enum WdPreferredWidthType
wdPreferredWidthPoints        =0x3        # from enum WdPreferredWidthType

wdPrintAutoTextEntries        =0x4        # from enum WdPrintOutItem
wdPrintComments               =0x2        # from enum WdPrintOutItem
wdPrintDocumentContent        =0x0        # from enum WdPrintOutItem
wdPrintDocumentWithMarkup     =0x7        # from enum WdPrintOutItem
wdPrintEnvelope               =0x6        # from enum WdPrintOutItem
wdPrintKeyAssignments         =0x5        # from enum WdPrintOutItem
wdPrintMarkup                 =0x2        # from enum WdPrintOutItem
wdPrintProperties             =0x1        # from enum WdPrintOutItem
wdPrintStyles                 =0x3        # from enum WdPrintOutItem

wdPrintAllPages               =0x0        # from enum WdPrintOutPages
wdPrintEvenPagesOnly          =0x2        # from enum WdPrintOutPages
wdPrintOddPagesOnly           =0x1        # from enum WdPrintOutPages
wdPrintAllDocument            =0x0        # from enum WdPrintOutRange
wdPrintCurrentPage            =0x2        # from enum WdPrintOutRange
wdPrintFromTo                 =0x3        # from enum WdPrintOutRange
wdPrintRangeOfPages           =0x4        # from enum WdPrintOutRange
wdPrintSelection              =0x1        # from enum WdPrintOutRange


wdLineBetweenColumnRectangle  =0x5        # from enum WdRectangleType
wdMarkupRectangle             =0x2        # from enum WdRectangleType
wdMarkupRectangleButton       =0x3        # from enum WdRectangleType
wdPageBorderRectangle         =0x4        # from enum WdRectangleType
wdSelection                   =0x6        # from enum WdRectangleType
wdShapeRectangle              =0x1        # from enum WdRectangleType
wdSystem                      =0x7        # from enum WdRectangleType
wdTextRectangle               =0x0        # from enum WdRectangleType

wdContentText                 =-1         # from enum WdReferenceKind
wdEndnoteNumber               =0x6        # from enum WdReferenceKind
wdEndnoteNumberFormatted      =0x11       # from enum WdReferenceKind
wdEntireCaption               =0x2        # from enum WdReferenceKind
wdFootnoteNumber              =0x5        # from enum WdReferenceKind
wdFootnoteNumberFormatted     =0x10       # from enum WdReferenceKind
wdNumberFullContext           =-4         # from enum WdReferenceKind
wdNumberNoContext             =-3         # from enum WdReferenceKind
wdNumberRelativeContext       =-2         # from enum WdReferenceKind
wdOnlyCaptionText             =0x4        # from enum WdReferenceKind
wdOnlyLabelAndNumber          =0x3        # from enum WdReferenceKind
wdPageNumber                  =0x7        # from enum WdReferenceKind
wdPosition                    =0xf        # from enum WdReferenceKind

wdRefTypeBookmark             =0x2        # from enum WdReferenceType
wdRefTypeEndnote              =0x4        # from enum WdReferenceType
wdRefTypeFootnote             =0x3        # from enum WdReferenceType
wdRefTypeHeading              =0x1        # from enum WdReferenceType
wdRefTypeNumberedItem         =0x0        # from enum WdReferenceType

wdRelativeHorizontalPositionCharacter=0x3       # from enum WdRelativeHorizontalPosition
wdRelativeHorizontalPositionColumn=0x2          # from enum WdRelativeHorizontalPosition
wdRelativeHorizontalPositionMargin=0x0          # from enum WdRelativeHorizontalPosition
wdRelativeHorizontalPositionPage=0x1            # from enum WdRelativeHorizontalPosition
wdRelativeVerticalPositionLine=0x3              # from enum WdRelativeVerticalPosition
wdRelativeVerticalPositionMargin=0x0            # from enum WdRelativeVerticalPosition
wdRelativeVerticalPositionPage=0x1              # from enum WdRelativeVerticalPosition
wdRelativeVerticalPositionParagraph=0x2         # from enum WdRelativeVerticalPosition

wdRelocateDown                =0x1        # from enum WdRelocate
wdRelocateUp                  =0x0        # from enum WdRelocate
wdReplaceAll                  =0x2        # from enum WdReplace
wdReplaceNone                 =0x0        # from enum WdReplace
wdReplaceOne                  =0x1        # from enum WdReplace

#TABLES:

wdAlignRowCenter              =0x1        # from enum WdRowAlignment
wdAlignRowLeft                =0x0        # from enum WdRowAlignment
wdAlignRowRight               =0x2        # from enum WdRowAlignment
wdRowHeightAtLeast            =0x1        # from enum WdRowHeightRule
wdRowHeightAuto               =0x0        # from enum WdRowHeightRule
wdRowHeightExactly            =0x2        # from enum WdRowHeightRule
wdAdjustFirstColumn           =0x2        # from enum WdRulerStyle
wdAdjustNone                  =0x0        # from enum WdRulerStyle
wdAdjustProportional          =0x1        # from enum WdRulerStyle
wdAdjustSameWidth             =0x3        # from enum WdRulerStyle


wdFormatDOSText               =0x4        # from enum WdSaveFormat
wdFormatDOSTextLineBreaks     =0x5        # from enum WdSaveFormat
wdFormatDocument              =0x0        # from enum WdSaveFormat
wdFormatEncodedText           =0x7        # from enum WdSaveFormat
wdFormatFilteredHTML          =0xa        # from enum WdSaveFormat
wdFormatHTML                  =0x8        # from enum WdSaveFormat
wdFormatRTF                   =0x6        # from enum WdSaveFormat
wdFormatTemplate              =0x1        # from enum WdSaveFormat
wdFormatText                  =0x2        # from enum WdSaveFormat
wdFormatTextLineBreaks        =0x3        # from enum WdSaveFormat
wdFormatUnicodeText           =0x7        # from enum WdSaveFormat
wdFormatWebArchive            =0x9        # from enum WdSaveFormat
wdFormatXML                   =0xb        # from enum WdSaveFormat
wdDoNotSaveChanges            =0x0        # from enum WdSaveOptions
wdPromptToSaveChanges         =-2         # from enum WdSaveOptions
wdSaveChanges                 =-1         # from enum WdSaveOptions

#SECTION

wdSectionContinuous           =0x0        # from enum WdSectionStart
wdSectionEvenPage             =0x3        # from enum WdSectionStart
wdSectionNewColumn            =0x1        # from enum WdSectionStart
wdSectionNewPage              =0x2        # from enum WdSectionStart
wdSectionOddPage              =0x4        # from enum WdSectionStart


wdSeekCurrentPageFooter       =0xa        # from enum WdSeekView
wdSeekCurrentPageHeader       =0x9        # from enum WdSeekView
wdSeekEndnotes                =0x8        # from enum WdSeekView
wdSeekEvenPagesFooter         =0x6        # from enum WdSeekView
wdSeekEvenPagesHeader         =0x3        # from enum WdSeekView
wdSeekFirstPageFooter         =0x5        # from enum WdSeekView
wdSeekFirstPageHeader         =0x2        # from enum WdSeekView
wdSeekFootnotes               =0x7        # from enum WdSeekView
wdSeekMainDocument            =0x0        # from enum WdSeekView
wdSeekPrimaryFooter           =0x4        # from enum WdSeekView
wdSeekPrimaryHeader           =0x1        # from enum WdSeekView

#SECTIONSELECTION

wdSelActive                   =0x8        # from enum WdSelectionFlags
wdSelAtEOL                    =0x2        # from enum WdSelectionFlags
wdSelOvertype                 =0x4        # from enum WdSelectionFlags
wdSelReplace                  =0x10       # from enum WdSelectionFlags
wdSelStartActive              =0x1        # from enum WdSelectionFlags


wdNoSelection                 =0x0        # from enum WdSelectionType
wdSelectionBlock              =0x6        # from enum WdSelectionType
wdSelectionColumn             =0x4        # from enum WdSelectionType
wdSelectionFrame              =0x3        # from enum WdSelectionType
wdSelectionIP                 =0x1        # from enum WdSelectionType
wdSelectionInlineShape        =0x7        # from enum WdSelectionType
wdSelectionNormal             =0x2        # from enum WdSelectionType
wdSelectionRow                =0x5        # from enum WdSelectionType
wdSelectionShape              =0x8        # from enum WdSelectionType
wdSeparatorColon              =0x2        # from enum WdSeparatorType
wdSeparatorEmDash             =0x3        # from enum WdSeparatorType
wdSeparatorEnDash             =0x4        # from enum WdSeparatorType
wdSeparatorHyphen             =0x0        # from enum WdSeparatorType
wdSeparatorPeriod             =0x1        # from enum WdSeparatorType

wdShapeBottom                 =-999997    # from enum WdShapePosition
wdShapeCenter                 =-999995    # from enum WdShapePosition
wdShapeInside                 =-999994    # from enum WdShapePosition
wdShapeLeft                   =-999998    # from enum WdShapePosition
wdShapeOutside                =-999993    # from enum WdShapePosition
wdShapeRight                  =-999996    # from enum WdShapePosition
wdShapeTop                    =-999999    # from enum WdShapePosition


wdSortOrderAscending          =0x0        # from enum WdSortOrder
wdSortOrderDescending         =0x1        # from enum WdSortOrder
wdSortSeparateByCommas        =0x1        # from enum WdSortSeparator
wdSortSeparateByDefaultTableSeparator=0x2        # from enum WdSortSeparator
wdSortSeparateByTabs          =0x0        # from enum WdSortSeparator


wdAlignTabBar                 =0x4        # from enum WdTabAlignment
wdAlignTabCenter              =0x1        # from enum WdTabAlignment
wdAlignTabDecimal             =0x3        # from enum WdTabAlignment
wdAlignTabLeft                =0x0        # from enum WdTabAlignment
wdAlignTabList                =0x6        # from enum WdTabAlignment
wdAlignTabRight               =0x2        # from enum WdTabAlignment

wdTabLeaderDashes             =0x2        # from enum WdTabLeader
wdTabLeaderDots               =0x1        # from enum WdTabLeader
wdTabLeaderHeavy              =0x4        # from enum WdTabLeader
wdTabLeaderLines              =0x3        # from enum WdTabLeader
wdTabLeaderMiddleDot          =0x5        # from enum WdTabLeader
wdTabLeaderSpaces             =0x0        # from enum WdTabLeader

#TABLE

wdTableDirectionLtr           =0x1        # from enum WdTableDirection
wdTableDirectionRtl           =0x0        # from enum WdTableDirection

wdSeparateByCommas            =0x2        # from enum WdTableFieldSeparator
wdSeparateByDefaultListSeparator=0x3        # from enum WdTableFieldSeparator
wdSeparateByParagraphs        =0x0        # from enum WdTableFieldSeparator
wdSeparateByTabs              =0x1        # from enum WdTableFieldSeparator

wdTableFormat3DEffects1       =0x20       # from enum WdTableFormat
wdTableFormat3DEffects2       =0x21       # from enum WdTableFormat
wdTableFormat3DEffects3       =0x22       # from enum WdTableFormat
wdTableFormatClassic1         =0x4        # from enum WdTableFormat
wdTableFormatClassic2         =0x5        # from enum WdTableFormat
wdTableFormatClassic3         =0x6        # from enum WdTableFormat
wdTableFormatClassic4         =0x7        # from enum WdTableFormat
wdTableFormatColorful1        =0x8        # from enum WdTableFormat
wdTableFormatColorful2        =0x9        # from enum WdTableFormat
wdTableFormatColorful3        =0xa        # from enum WdTableFormat
wdTableFormatColumns1         =0xb        # from enum WdTableFormat
wdTableFormatColumns2         =0xc        # from enum WdTableFormat
wdTableFormatColumns3         =0xd        # from enum WdTableFormat
wdTableFormatColumns4         =0xe        # from enum WdTableFormat
wdTableFormatColumns5         =0xf        # from enum WdTableFormat
wdTableFormatContemporary     =0x23       # from enum WdTableFormat
wdTableFormatElegant          =0x24       # from enum WdTableFormat
wdTableFormatGrid1            =0x10       # from enum WdTableFormat
wdTableFormatGrid2            =0x11       # from enum WdTableFormat
wdTableFormatGrid3            =0x12       # from enum WdTableFormat
wdTableFormatGrid4            =0x13       # from enum WdTableFormat
wdTableFormatGrid5            =0x14       # from enum WdTableFormat
wdTableFormatGrid6            =0x15       # from enum WdTableFormat
wdTableFormatGrid7            =0x16       # from enum WdTableFormat
wdTableFormatGrid8            =0x17       # from enum WdTableFormat
wdTableFormatList1            =0x18       # from enum WdTableFormat
wdTableFormatList2            =0x19       # from enum WdTableFormat
wdTableFormatList3            =0x1a       # from enum WdTableFormat
wdTableFormatList4            =0x1b       # from enum WdTableFormat
wdTableFormatList5            =0x1c       # from enum WdTableFormat
wdTableFormatList6            =0x1d       # from enum WdTableFormat
wdTableFormatList7            =0x1e       # from enum WdTableFormat
wdTableFormatList8            =0x1f       # from enum WdTableFormat
wdTableFormatNone             =0x0        # from enum WdTableFormat
wdTableFormatProfessional     =0x25       # from enum WdTableFormat
wdTableFormatSimple1          =0x1        # from enum WdTableFormat
wdTableFormatSimple2          =0x2        # from enum WdTableFormat
wdTableFormatSimple3          =0x3        # from enum WdTableFormat
wdTableFormatSubtle1          =0x26       # from enum WdTableFormat
wdTableFormatSubtle2          =0x27       # from enum WdTableFormat
wdTableFormatWeb1             =0x28       # from enum WdTableFormat
wdTableFormatWeb2             =0x29       # from enum WdTableFormat
wdTableFormatWeb3             =0x2a       # from enum WdTableFormat

wdTableFormatApplyAutoFit     =0x10       # from enum WdTableFormatApply
wdTableFormatApplyBorders     =0x1        # from enum WdTableFormatApply
wdTableFormatApplyColor       =0x8        # from enum WdTableFormatApply
wdTableFormatApplyFirstColumn =0x80       # from enum WdTableFormatApply
wdTableFormatApplyFont        =0x4        # from enum WdTableFormatApply
wdTableFormatApplyHeadingRows =0x20       # from enum WdTableFormatApply
wdTableFormatApplyLastColumn  =0x100      # from enum WdTableFormatApply
wdTableFormatApplyLastRow     =0x40       # from enum WdTableFormatApply
wdTableFormatApplyShading     =0x2        # from enum WdTableFormatApply

wdTableBottom                 =-999997    # from enum WdTablePosition
wdTableCenter                 =-999995    # from enum WdTablePosition
wdTableInside                 =-999994    # from enum WdTablePosition
wdTableLeft                   =-999998    # from enum WdTablePosition
wdTableOutside                =-999993    # from enum WdTablePosition
wdTableRight                  =-999996    # from enum WdTablePosition
wdTableTop                    =-999999    # from enum WdTablePosition


wdAttachedTemplate            =0x2        # from enum WdTemplateType
wdGlobalTemplate              =0x1        # from enum WdTemplateType
wdNormalTemplate              =0x0        # from enum WdTemplateType

wdCalculationText             =0x5        # from enum WdTextFormFieldType
wdCurrentDateText             =0x3        # from enum WdTextFormFieldType
wdCurrentTimeText             =0x4        # from enum WdTextFormFieldType
wdDateText                    =0x2        # from enum WdTextFormFieldType
wdNumberText                  =0x1        # from enum WdTextFormFieldType
wdRegularText                 =0x0        # from enum WdTextFormFieldType

wdTextOrientationDownward     =0x3        # from enum WdTextOrientation
wdTextOrientationHorizontal   =0x0        # from enum WdTextOrientation
wdTextOrientationUpward       =0x2        # from enum WdTextOrientation


wdTOCClassic                  =0x1        # from enum WdTocFormat
wdTOCDistinctive              =0x2        # from enum WdTocFormat
wdTOCFancy                    =0x3        # from enum WdTocFormat
wdTOCFormal                   =0x5        # from enum WdTocFormat
wdTOCModern                   =0x4        # from enum WdTocFormat
wdTOCSimple                   =0x6        # from enum WdTocFormat
wdTOCTemplate                 =0x0        # from enum WdTocFormat
wdTOFCentered                 =0x3        # from enum WdTofFormat
wdTOFClassic                  =0x1        # from enum WdTofFormat
wdTOFDistinctive              =0x2        # from enum WdTofFormat
wdTOFFormal                   =0x4        # from enum WdTofFormat
wdTOFSimple                   =0x5        # from enum WdTofFormat
wdTOFTemplate                 =0x0        # from enum WdTofFormat


wdCell                        =0xc        # from enum WdUnits
wdCharacter                   =0x1        # from enum WdUnits
wdCharacterFormatting         =0xd        # from enum WdUnits
wdColumn                      =0x9        # from enum WdUnits
wdItem                        =0x10       # from enum WdUnits
wdLine                        =0x5        # from enum WdUnits
wdParagraph                   =0x4        # from enum WdUnits
wdParagraphFormatting         =0xe        # from enum WdUnits
wdRow                         =0xa        # from enum WdUnits
wdScreen                      =0x7        # from enum WdUnits
wdSection                     =0x8        # from enum WdUnits
wdSentence                    =0x3        # from enum WdUnits
wdStory                       =0x6        # from enum WdUnits
wdTable                       =0xf        # from enum WdUnits
wdWindow                      =0xb        # from enum WdUnits
wdWord                        =0x2        # from enum WdUnits

wdMove                        =0x0        # The selection is collapsed to an insertion point and moved to the end of the specified unit. Default.
wdExtend                      =0x1        # The end of the selection is extended to the end of the specified unit.

wdFormattingFromCurrent       =0x0        # from enum WdUseFormattingFrom
wdFormattingFromPrompt        =0x2        # from enum WdUseFormattingFrom
wdFormattingFromSelected      =0x1        # from enum WdUseFormattingFrom

wdAlignVerticalBottom         =0x3        # from enum WdVerticalAlignment
wdAlignVerticalCenter         =0x1        # from enum WdVerticalAlignment
wdAlignVerticalJustify        =0x2        # from enum WdVerticalAlignment
wdAlignVerticalTop            =0x0        # from enum WdVerticalAlignment



wdVisualSelectionBlock        =0x0        # from enum WdVisualSelection
wdVisualSelectionContinuous   =0x1        # from enum WdVisualSelection


wdWrapBoth                    =0x0        # from enum WdWrapSideType
wdWrapLargest                 =0x3        # from enum WdWrapSideType
wdWrapLeft                    =0x1        # from enum WdWrapSideType
wdWrapRight                   =0x2        # from enum WdWrapSideType

wdWrapInline                  =0x7        # from enum WdWrapType
wdWrapNone                    =0x3        # from enum WdWrapType
wdWrapSquare                  =0x0        # from enum WdWrapType
wdWrapThrough                 =0x2        # from enum WdWrapType
wdWrapTight                   =0x1        # from enum WdWrapType
wdWrapTopBottom               =0x4        # from enum WdWrapType

wdWrapMergeBehind             =0x3        # from enum WdWrapTypeMerged
wdWrapMergeFront              =0x4        # from enum WdWrapTypeMerged
wdWrapMergeInline             =0x0        # from enum WdWrapTypeMerged
wdWrapMergeSquare             =0x1        # from enum WdWrapTypeMerged
wdWrapMergeThrough            =0x5        # from enum WdWrapTypeMerged
wdWrapMergeTight              =0x2        # from enum WdWrapTypeMerged
wdWrapMergeTopBottom          =0x6        # from enum WdWrapTypeMerged

xlUserResolution              =0x1        # from enum ConflictResolution 
xlLocalSessionChanges         =0x2        # from enum ConflictResolution 



# Event Handlers
# If you create handlers, they should have the following prototypes:
#	def OnQuit(self):
#	def OnDocumentBeforePrint(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnWindowBeforeRightClick(self, Sel=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnStartup(self):
#	def OnNewDocument(self, Doc=defaultNamedNotOptArg):
#	def OnDocumentBeforeSave(self, Doc=defaultNamedNotOptArg, SaveAsUI=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnDocumentChange(self):
#	def OnWindowBeforeDoubleClick(self, Sel=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnWindowActivate(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnAddRef(self):
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnWindowDeactivate(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnDocumentBeforeClose(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnRelease(self):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, rgdispid=pythoncom.Missing):
#	def OnDocumentOpen(self, Doc=defaultNamedNotOptArg):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnWindowSelectionChange(self, Sel=defaultNamedNotOptArg):


# Event Handlers
# If you create handlers, they should have the following prototypes:
#	def OnQuit(self):
#	def OnMailMergeBeforeMerge(self, Doc=defaultNamedNotOptArg, StartRecord=defaultNamedNotOptArg, EndRecord=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnMailMergeBeforeRecordMerge(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnDocumentBeforeClose(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnMailMergeDataSourceValidate(self, Doc=defaultNamedNotOptArg, Handled=defaultNamedNotOptArg):
#	def OnWindowSize(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnDocumentOpen(self, Doc=defaultNamedNotOptArg):
#	def OnMailMergeAfterRecordMerge(self, Doc=defaultNamedNotOptArg):
#	def OnWindowBeforeDoubleClick(self, Sel=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnDocumentBeforeSave(self, Doc=defaultNamedNotOptArg, SaveAsUI=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnWindowBeforeRightClick(self, Sel=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnEPostageInsert(self, Doc=defaultNamedNotOptArg):
#	def OnWindowDeactivate(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnRelease(self):
#	def OnEPostagePropertyDialog(self, Doc=defaultNamedNotOptArg):
#	def OnMailMergeAfterMerge(self, Doc=defaultNamedNotOptArg, DocResult=defaultNamedNotOptArg):
#	def OnDocumentBeforePrint(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnNewDocument(self, Doc=defaultNamedNotOptArg):
#	def OnStartup(self):
#	def OnDocumentChange(self):
#	def OnMailMergeWizardSendToCustom(self, Doc=defaultNamedNotOptArg):
#	def OnWindowActivate(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnMailMergeDataSourceLoad(self, Doc=defaultNamedNotOptArg):
#	def OnWindowSelectionChange(self, Sel=defaultNamedNotOptArg):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, rgdispid=pythoncom.Missing):
#	def OnMailMergeWizardStateChange(self, Doc=defaultNamedNotOptArg, FromState=defaultNamedNotOptArg, ToState=defaultNamedNotOptArg, Handled=defaultNamedNotOptArg):


# Event Handlers
# If you create handlers, they should have the following prototypes:
#	def OnQuit(self):
#	def OnMailMergeBeforeMerge(self, Doc=defaultNamedNotOptArg, StartRecord=defaultNamedNotOptArg, EndRecord=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnMailMergeBeforeRecordMerge(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnDocumentBeforeClose(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnMailMergeDataSourceValidate(self, Doc=defaultNamedNotOptArg, Handled=defaultNamedNotOptArg):
#	def OnWindowSize(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnEPostageInsertEx(self, Doc=defaultNamedNotOptArg, cpDeliveryAddrStart=defaultNamedNotOptArg, cpDeliveryAddrEnd=defaultNamedNotOptArg, cpReturnAddrStart=defaultNamedNotOptArg, cpReturnAddrEnd=defaultNamedNotOptArg, xaWidth=defaultNamedNotOptArg, yaHeight=defaultNamedNotOptArg, bstrPrinterName=defaultNamedNotOptArg, bstrPaperFeed=defaultNamedNotOptArg, fPrint=defaultNamedNotOptArg, fCancel=defaultNamedNotOptArg):
#	def OnXMLSelectionChange(self, Sel=defaultNamedNotOptArg, OldXMLNode=defaultNamedNotOptArg, NewXMLNode=defaultNamedNotOptArg, Reason=defaultNamedNotOptArg):
#	def OnDocumentOpen(self, Doc=defaultNamedNotOptArg):
#	def OnMailMergeAfterRecordMerge(self, Doc=defaultNamedNotOptArg):
#	def OnWindowBeforeDoubleClick(self, Sel=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnDocumentSync(self, Doc=defaultNamedNotOptArg, SyncEventType=defaultNamedNotOptArg):
#	def OnDocumentBeforeSave(self, Doc=defaultNamedNotOptArg, SaveAsUI=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnWindowBeforeRightClick(self, Sel=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnEPostageInsert(self, Doc=defaultNamedNotOptArg):
#	def OnWindowDeactivate(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnRelease(self):
#	def OnEPostagePropertyDialog(self, Doc=defaultNamedNotOptArg):
#	def OnMailMergeAfterMerge(self, Doc=defaultNamedNotOptArg, DocResult=defaultNamedNotOptArg):
#	def OnDocumentBeforePrint(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnNewDocument(self, Doc=defaultNamedNotOptArg):
#	def OnStartup(self):
#	def OnDocumentChange(self):
#	def OnMailMergeWizardSendToCustom(self, Doc=defaultNamedNotOptArg):
#	def OnWindowActivate(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnMailMergeDataSourceLoad(self, Doc=defaultNamedNotOptArg):
#	def OnWindowSelectionChange(self, Sel=defaultNamedNotOptArg):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, rgdispid=pythoncom.Missing):
#	def OnMailMergeWizardStateChange(self, Doc=defaultNamedNotOptArg, FromState=defaultNamedNotOptArg, ToState=defaultNamedNotOptArg, Handled=defaultNamedNotOptArg):
#	def OnXMLValidationError(self, XMLNode=defaultNamedNotOptArg):

