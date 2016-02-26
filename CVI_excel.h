/**************************************************************************/
/* LabWindows/CVI User Interface Resource (UIR) Include File              */
/* Copyright (c) National Instruments 2016. All Rights Reserved.          */
/*                                                                        */
/* WARNING: Do not add to, delete from, or otherwise modify the contents  */
/*          of this include file.                                         */
/**************************************************************************/

#include <userint.h>

#ifdef __cplusplus
    extern "C" {
#endif

     /* Panels and Controls: */

#define  PANEL                            1       /* callback function: QuitCallback */
#define  PANEL_BTN_LAUNCH_EXCEL           2       /* control type: command, callback function: launch_excel */
#define  PANEL_BTN_SHUT_DOWN_EXCEL        3       /* control type: command, callback function: shut_down_excel */
#define  PANEL_BTN_OPEN_EXCEL             4       /* control type: command, callback function: open_excel */
#define  PANEL_TEXTBOX_POD                5       /* control type: textBox, callback function: (none) */
#define  PANEL_BTN_PROCESS_SHEET          6       /* control type: command, callback function: process_sheet */
#define  PANEL_NUMERIC_ROW                7       /* control type: numeric, callback function: (none) */
#define  PANEL_NUMERIC_COLUMN             8       /* control type: numeric, callback function: (none) */
#define  PANEL_DECORATION                 9       /* control type: deco, callback function: (none) */


     /* Control Arrays: */

          /* (no control arrays in the resource file) */


     /* Menu Bars, Menus, and Menu Items: */

          /* (no menu bars in the resource file) */


     /* Callback Prototypes: */

int  CVICALLBACK launch_excel(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK open_excel(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK process_sheet(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK QuitCallback(int panel, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK shut_down_excel(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);


#ifdef __cplusplus
    }
#endif
