Attribute VB_Name = "modConstants"
Option Explicit
'================================================================================
' MODULE: modConstants
' PURPOSE: Central location for stable formatting/layout constants used across
'          formatting and pipeline modules. These were previously implicit or
'          removed during cleanup – reintroducing them avoids magic numbers and
'          unresolved identifiers.
' NOTE:    If a future configuration system exposes these, keep the constants
'          as defaults and let runtime override vars shadow them.
'================================================================================

' ---- Font & Typography ----
Public Const STANDARD_FONT As String = "Times New Roman"
Public Const STANDARD_FONT_SIZE As Long = 12
Public Const FOOTER_FONT_SIZE As Long = 10

' When using wdLineSpaceMultiple Word expects: desired_lines * 12 (points).
' For 1.15 lines (common in legislative docs): 1.15 * 12 = 13.8 (rounded 14).
Public Const LINE_SPACING As Long = 14

' ---- Page Layout (centimeters) ----
Public Const TOP_MARGIN_CM As Double = 2.0
Public Const BOTTOM_MARGIN_CM As Double = 2.0
Public Const LEFT_MARGIN_CM As Double = 2.0
Public Const RIGHT_MARGIN_CM As Double = 2.0
Public Const HEADER_DISTANCE_CM As Double = 1.0
Public Const FOOTER_DISTANCE_CM As Double = 1.0

' ---- Header Image ----
Public Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 6.5
Public Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.35  ' width * ratio = height
Public Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 1.0

' ---- Misc ----
Public Const VERSION_STRING As String = "1.0.0-Beta3"

' End of file.
