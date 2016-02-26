#include <cvirte.h>	
#include <userint.h>
#include <cviauto.h>
#include <utility.h> 

#include "CVI_excel.h"
#include "excel2000.h"
#include "excelreport.h"
#include "toolbox.h" 
//----------------------------------------------------------------------------
// Defines
//----------------------------------------------------------------------------
// errChk - Macro in toolbox.h that will goto label "Error:"
//          if error returned is less than zero.
#define caErrChk errChk  
#define APP_AUTOMATION_ERR "Error:  Microsoft Excel Automation"
#define APP_WARNING "Warning"
#define EXCEL_ARRAY_OF_CELLS "A1:M1"    // A4:M4
#define ROWS 10
#define ROW_STRING 3   // 0
#define COLUMNS 27
#define LAUNCHERR "\
An error occurred trying to launch Excel 2000 through its automation interface.\n\n\
Ensure that Excel is installed and that you can launch it manually. If errors\n\
persist, try to launch Excel manually and use the CONNECT button instead."


#define APP_AUTOMATION_ERR "Error:  Microsoft Excel Automation"

#define APP_WARNING "Warning"

#define BUF_LEN 6
#define MANUALJIG_1 32
#define MANUALJIG_2 31  
#define POD1 1
#define POD2 2
#define POD3 3
#define POD4 4
#define POD5 5
#define JIG1 1 
#define JIG2 2
#define JIG3 3
#define JIG4 4
#define PASS_WIFI 10
#define FAIL_WIFI 12
#define PART_FAIL_WIFI 13


#define MAX_ROW 4
#define MAX_COL 13                  // was 5
#define MAX_POD 7
#define MAX_JIG 4
#define MAX_TYPE 3

//---------
static int panelHandle;  

static int excelLaunched = 0;
static int appVisible = 1;

//-------


static HRESULT status;

static ExcelObj_App               ExcelAppHandle = 0;       
static ExcelObj_Workbooks         ExcelWorkbooksHandle = 0; 
static ExcelObj_Workbook          ExcelWorkbookHandle = 0;  
static ExcelObj_Sheets            ExcelSheetsHandle = 0;    
static ExcelObj_Worksheet         ExcelWorksheetHandle = 0; 
static ExcelObj_Range             ExcelRangeHandle = 0;     
static ExcelObj_ChartObject       ExcelChartObjHandle = 0;
static ExcelObj_Chart             ExcelChartHandle = 0;
static ExcelObj_ChartGroup        ExcelChartsHandle = 0;

																						
static ERRORINFO ErrorInfo;
static VARIANT MyVariant;
static LPDISPATCH MyDispatch;
static VARIANT MyCellRangeV;
static CAObjHandle chartHandle = 0;

int ***auto_Pod_Pi;

char *autopod_Pi[] ={"auto_pod_1","auto_pod_2","auto_pod_3","auto_pod_4","auto_pod_5","manual_Jig_1","manual_Jig_2"};

char *jig_pi[]={"          jig_1:  ","          jig_2:  ","          jig_3:  ","          jig_4:  "};


char *wifi_Type_Test[]={"WIFI_LINK_QUALITY  : ","WIFI_PING                  : ","WIFI_SIGNAL_LEVEL : "};
//***************************************************************************************
// Prototypes
//***************************************************************************************

static HRESULT SaveDocument (CAObjHandle ExcelWorksheetHandle, char *fileName);
HRESULT ClearObjHandle(CAObjHandle *objHandle);

static int  ShutdownExcel(void);
static void ReportAppAutomationError (HRESULT hr);
static void InitVariables(void);
static int  UpdateUIRDimming(int panel);
void get_Value_row_col(int *row , int *col);
int process_Value_row(int buf[MAX_TYPE][BUF_LEN],int row , int col, int **outPOD[MAX_POD]) ;
int init_Globla_var(void);
void copy_Pod(int **auto_Pod[MAX_POD] ,int **auto_Pod_Out[MAX_POD], int row, int col, int num_Pod);
int display_Result_Pi(int **auto_Pod[MAX_POD], int row , int col , int num_Pod);
void free_Global_Var(void);
HRESULT ReadDataFromExcel(int **auto_Pod[MAX_POD]);

int main (int argc, char *argv[])
{
	if (InitCVIRTE (0, argv, 0) == 0)
		return -1;	/* out of memory */
	
	CA_InitActiveXThreadStyleForCurrentThread (0, COINIT_APARTMENTTHREADED);
	
	init_Globla_var();
	
    SetSleepPolicy (VAL_SLEEP_MORE);
	
	if ((panelHandle = LoadPanel (0, "CVI_excel.uir", PANEL)) < 0)
		return -1;
	DisplayPanel (panelHandle);
	RunUserInterface ();
	DiscardPanel (panelHandle);
	free_Global_Var();
	return 0;
}

//********************************** ALL CALLBACK BELOW **********************************
//
//
//****************************************************************************************


int CVICALLBACK launch_excel (int panel, int control, int event,
		void *callbackData, int eventData1, int eventData2)
{
	HRESULT error = 0;
	
	switch (event)
	{
		case EVENT_COMMIT:
			
			 SetWaitCursor (1);
			 error = Excel_NewApp (NULL, 1, LOCALE_NEUTRAL, 0, &ExcelAppHandle);
			 SetWaitCursor (0);
             if (error<0) 
             {
                MessagePopup (APP_AUTOMATION_ERR, LAUNCHERR);
                error = 0;
             }
             // Make App Visible
             error = Excel_SetProperty (ExcelAppHandle, NULL, Excel_AppVisible, CAVT_BOOL, appVisible?VTRUE:VFALSE);
             if (error<0)
			 {
				 ReportAppAutomationError (error);
			 }
			  
			 MakeApplicationActive ();
			 
			 break;
		case EVENT_LOST_FOCUS:

			break;
		case EVENT_DISCARD:

			break;
	}
	return 0;
}

int CVICALLBACK shut_down_excel (int panel, int control, int event,
		void *callbackData, int eventData1, int eventData2)
{
	switch (event)
	{
		case EVENT_COMMIT:
			 SetWaitCursor (1);
             ShutdownExcel();  
             SetWaitCursor (0);
			break;
	}
	return 0;
}

int CVICALLBACK outPOD (int panel, int control, int event,
		void *callbackData, int eventData1, int eventData2)
{
	switch (event)
	{
		case EVENT_COMMIT:

			break;
		case EVENT_LOST_FOCUS:

			break;
		case EVENT_DISCARD:

			break;
	}
	return 0;
}

int CVICALLBACK QuitCallback (int panel, int event, void *callbackData,
		int eventData1, int eventData2)
{
	switch (event)
	{
		case EVENT_GOT_FOCUS:

			break;
		case EVENT_LOST_FOCUS:

			break;
		case EVENT_CLOSE:
			  QuitUserInterface(0);
			break;
		case EVENT_PANEL_SIZING:

			break;
	}
	return 0;
}

int CVICALLBACK open_excel (int panel, int control, int event,   // I AM HERE 11-02-2016
		void *callbackData, int eventData1, int eventData2)
{
	int ret;
	
	HRESULT error = 0;
    char fileName[MAX_PATHNAME_LEN];
	
	switch (event)
	{
		case EVENT_COMMIT:
			if (!ExcelWorkbooksHandle)
			{
			   error = Excel_GetProperty (ExcelAppHandle, NULL, Excel_AppWorkbooks, CAVT_OBJHANDLE, &ExcelWorkbooksHandle);
			    if (error<0)
				{
				    ClearObjHandle (&ExcelWorksheetHandle);
					ClearObjHandle (&ExcelSheetsHandle);
					ClearObjHandle (&ExcelWorkbookHandle);
					ClearObjHandle (&ExcelWorkbooksHandle);
        
                    if (error < 0)
					{
					  ReportAppAutomationError (error); 
					}
				}
				GetProjectDir (fileName);
				strcat(fileName, "\\log_browser_output2.xlsx");
				error = Excel_WorkbooksOpen (ExcelWorkbooksHandle, NULL, fileName, CA_DEFAULT_VAL,
                                             CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                             CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                             CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                             CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                             CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                             CA_DEFAULT_VAL, &ExcelWorkbookHandle);
			    if (error<0)
				{
				    ClearObjHandle (&ExcelWorksheetHandle);
					ClearObjHandle (&ExcelSheetsHandle);
					ClearObjHandle (&ExcelWorkbookHandle);
					ClearObjHandle (&ExcelWorkbooksHandle);
        
                    if (error < 0)
					{
					  ReportAppAutomationError (error); 
					}
				}
				error = Excel_GetProperty (ExcelAppHandle, NULL, Excel_AppSheets,CAVT_OBJHANDLE, &ExcelSheetsHandle);
			    if (error<0)
				{
				    ClearObjHandle (&ExcelWorksheetHandle);
					ClearObjHandle (&ExcelSheetsHandle);
					ClearObjHandle (&ExcelWorkbookHandle);
					ClearObjHandle (&ExcelWorkbooksHandle);
        
                    if (error < 0)
					{
					  ReportAppAutomationError (error); 
					}
				}
				 error = Excel_SheetsItem (ExcelSheetsHandle, NULL, CA_VariantInt(1), &ExcelWorksheetHandle);
			    if (error<0)
				{
				    ClearObjHandle (&ExcelWorksheetHandle);
					ClearObjHandle (&ExcelSheetsHandle);
					ClearObjHandle (&ExcelWorkbookHandle);
					ClearObjHandle (&ExcelWorkbooksHandle);
        
                    if (error < 0)
					{
					  ReportAppAutomationError (error); 
					}
				}
				  // Make First Sheet Active - should already be active    
                error = Excel_WorksheetActivate (ExcelWorksheetHandle, NULL);
			    if (error<0)
				{
				    ClearObjHandle (&ExcelWorksheetHandle);
					ClearObjHandle (&ExcelSheetsHandle);
					ClearObjHandle (&ExcelWorkbookHandle);
					ClearObjHandle (&ExcelWorkbooksHandle);
        
                    if (error < 0)
					{
					  ReportAppAutomationError (error); 
					}
				}
				
			}
			else
			{
                MessagePopup(APP_WARNING, "Document already open");
			}

			break;
		case EVENT_LOST_FOCUS:

			break;
		case EVENT_DISCARD:

			break;
	}
	return 0;
}

int CVICALLBACK process_sheet (int panel, int control, int event,
		void *callbackData, int eventData1, int eventData2)
{
	int ret=0; 
	switch (event)
	{
		case EVENT_COMMIT:
			  ReadDataFromExcel(auto_Pod_Pi);
			  ret = display_Result_Pi(auto_Pod_Pi, MAX_ROW,MAX_COL ,MAX_POD);
			break;
	}
	return 0;
}


//============================ END CALLBACK ==================================


//======================== ===================================================
// shut down excel 
//============================================================================

static int ShutdownExcel(void) 
{
    HRESULT error = 0;

    ClearObjHandle (&ExcelRangeHandle);
    ClearObjHandle (&ExcelWorksheetHandle);
    ClearObjHandle (&ExcelSheetsHandle);
    
    if (ExcelWorkbookHandle) 
    {
        // Close workbook without saving
        error = Excel_WorkbookClose (ExcelWorkbookHandle, NULL, CA_VariantBool (VFALSE), 
            CA_DEFAULT_VAL, CA_VariantBool (VFALSE));
        if (error < 0)
            goto Error;
        
        ClearObjHandle (&ExcelWorkbookHandle);
    }
    
    ClearObjHandle (&ExcelWorkbooksHandle);
        
    if (ExcelAppHandle)
    {   
        if (excelLaunched) 
        {
            // Quit the Application
            error = Excel_AppQuit (ExcelAppHandle, &ErrorInfo);
            if (error < 0) goto Error;
        }
        
        ClearObjHandle (&ExcelAppHandle);
    } 
    
    return 0;   
Error:    
    if (error < 0)
        ReportAppAutomationError (error);
        
    return error;                    
}


//===========================================================================
//  clearObj handled
//===========================================================================
HRESULT ClearObjHandle(CAObjHandle *objHandle)
{
    HRESULT error = 0;
    if ((objHandle) && (*objHandle))
    {
        error = CA_DiscardObjHandle (*objHandle);
        *objHandle = 0;
    }
    return error;    
}    

//----------------------------------------------------------------------------
// ReportWordAutomationError
//----------------------------------------------------------------------------
static void ReportAppAutomationError (HRESULT hr)
{
    char errorBuf[256];
    
    if (hr < 0) {
        CA_GetAutomationErrorString (hr, errorBuf, sizeof (errorBuf));
        MessagePopup (APP_AUTOMATION_ERR, errorBuf);
    }
    return;
}

//=================================================================================
//
//=================================================================================
HRESULT ReadDataFromExcel(int **auto_Pod[MAX_POD])
{

    HRESULT error = 0;
    int i;
	int j;
	int k;
	int rowValue;
	int colValue; 
	int wifi_link_qual_col=0;
	int wifi_ping_col=0;
	int wifi_signallevel_col=0; 
	int unit_col[MAX_TYPE]={0,0,0};
	int status_col[MAX_TYPE]={0,0,0};
	int jig_id_col[MAX_TYPE]={0,0,0};
	int jig_mode_col[MAX_TYPE]={0,0,0};
	int group_col[MAX_TYPE]={0,0,0};
	int num_col[MAX_TYPE]  ={0,0,0};
	int trace_row=0;
	int numb_board_pass=0; 
	int outPOD_number=0;
	int buffer_line[MAX_TYPE][BUF_LEN]={{0,0,0,0,0,0},{0,0,0,0,0,0},{0,0,0,0,0,0}}; 
	int number_pass=0;
	int number_fail=0;
	int number_Part_fail=0;
	int test_Type=0; 
    VARIANT *vArray = NULL;
    size_t dim1Size, dim2Size;
    double d;
	DATE date;
	char *str;
	char *strtest;
	char *str_status;
	int  flag_process_data=0;
	int ret=0; 
	double   int_numb;
	
	int ***buf_Pod ;                                   
	
	//--- init buf_Pod malloc memory ---
	
	buf_Pod = (int***)malloc(sizeof(int**)*MAX_POD);
	
	if(buf_Pod == NULL)
	{
	  return -1; 
	}
	
	for(i=0 ; i<MAX_POD; i++)
	{
	   buf_Pod[i]=(int**)malloc(sizeof(int*)*MAX_ROW);
	   
	   if(buf_Pod[i]==NULL)
	   {
	     return -1;
	   }
	   
	   for(j=0 ; j<MAX_ROW; j++)
	   {
	      buf_Pod[i][j]= (int*)malloc(sizeof(int)*MAX_COL);
		  
		  if(buf_Pod[i][j]==NULL)
		  {
		    return -1; 
		  }
		  
		  for(k=0 ; k<MAX_COL; k++)
		  {
		    buf_Pod[i][j][k]=0;
		  }
	   
	   }
	
	}
	
    ExcelObj_Range ExcelSingleCellRangeHandle = 0;
	
	get_Value_row_col(&rowValue ,&colValue); 
	
    SetWaitCursor (1);
    
    // Open new Range for Worksheet
    error = CA_VariantSetCString (&MyCellRangeV, EXCEL_ARRAY_OF_CELLS);
    error = Excel_WorksheetRange (ExcelWorksheetHandle, NULL, MyCellRangeV, CA_DEFAULT_VAL, &ExcelRangeHandle);
    CA_VariantClear(&MyCellRangeV);
    if (error<0) goto Error;

    // Make range Active    
    error = Excel_RangeActivate (ExcelRangeHandle, &ErrorInfo, NULL);
    if (error<0) goto Error;

    //----------------------------------------------------------------
    // 1) Get each cell value in Range one at a time using an offset 
    //    from the range's top left cell
    //----------------------------------------------------------------
    SetStdioWindowVisibility (1);
	
    for (i=0;i<rowValue;i++)
    {
	   
	   if( i== ROW_STRING)
	   {
		   
         for (j=0;j<colValue;j++)
         {
            error = Excel_RangeGetItem (ExcelRangeHandle, &ErrorInfo, CA_VariantInt (i+1), CA_VariantInt (j+1), &MyVariant);
            if (error<0) goto Error;
            error = CA_VariantGetDispatch (&MyVariant, &MyDispatch);
            if (error<0) goto Error;
            
         
            error = CA_CreateObjHandleFromIDispatch (MyDispatch, 0, &ExcelSingleCellRangeHandle);
            if (error<0) goto Error;
            
        
            error = Excel_GetProperty (ExcelSingleCellRangeHandle, &ErrorInfo, Excel_RangeValue2, CAVT_VARIANT, &MyVariant);
            if (error<0) goto Error;
			
			if(CA_VariantHasCString(&MyVariant) == 1)
			{
            	error = CA_VariantGetCString(&MyVariant, &str);
				if (error<0) goto Error;
				CA_VariantClear(&MyVariant);
				
				if(strcmp(str,"WIFI_LINKQUALITY")==0)
				{
					wifi_link_qual_col=j+1;
					
					buffer_line[0][2]=1;
					
					test_Type=1;
				}
				
				else if(strcmp(str,"WIFI_PING")==0)
				{
					wifi_ping_col = j+1;
					
					buffer_line[1][3]=2; 
					
					test_Type=2;
				
				}
				else if(strcmp(str,"WIFI_SIGNALLEVEL")==0)
				{
					wifi_signallevel_col= j+1;
					
					buffer_line[2][4]=3;
					
					test_Type=3; 
				}
				else if(strcmp(str,"UNITS")==0)
				{
					if(test_Type == 1)
					{
					  unit_col[0] = j+1; 
					}
					else if(test_Type == 2)
					{
					  unit_col[1] = j+1; 
					}
					else if(test_Type == 3)
					{
					  unit_col[2] = j+1; 
					}
					
				}
				else if(strcmp(str,"STATUS")==0)
				{
					
					if(test_Type == 1)
					{
					  status_col[0] = j+1; 
					}
					else if(test_Type == 2)
					{
					  status_col[1] = j+1; 
					}
					else if(test_Type == 3)
					{
					  status_col[2] = j+1; 
					}
					
				    str="";
				}
				else if(strcmp(str,"JIG_ID")==0)
				{
					if(test_Type == 1)
					{
					  jig_id_col[0] = j+1; 
					}
					else if(test_Type == 2)
					{
					  jig_id_col[1] = j+1; 
					}
					else if(test_Type == 3)
					{
					  jig_id_col[2] = j+1; 
					}
					
	
				}
				else if(strcmp(str,"JIG_MODE")==0)
				{
					
					if(test_Type == 1)
					{
					  jig_mode_col[0] = j+1; 
					}
					else if(test_Type == 2)
					{
					  jig_mode_col[1] = j+1; 
					}
					else if(test_Type == 3)
					{
					  jig_mode_col[2] = j+1; 
					}
					
				}
				
				else if(strcmp(str,"GROUP")==0)
				{
					
					if(test_Type == 1)
					{
					  group_col[0] = j+1; 
					}
					else if(test_Type == 2)
					{
					  group_col[1] = j+1; 
					}
					else if(test_Type == 3)
					{
					  group_col[2] = j+1; 
					}
					
				 
				}
				else if(strcmp(str,"NUM")==0)
				{
					if(test_Type == 1)
					{
					  num_col[0] = j+1; 
					}
					else if(test_Type == 2)
					{
					  num_col[1] = j+1; 
					}
					else if(test_Type == 3)
					{
					  num_col[2] = j+1; 
					}
					 
				}
				
		 	 }
		 }
	   }
	   else
	   {
		 flag_process_data = 1;
		 
		 test_Type = 0; 
		 
	     for (j=1;j<colValue;j++)
	     {
            error = Excel_RangeGetItem (ExcelRangeHandle, &ErrorInfo, CA_VariantInt (i+1), CA_VariantInt (j), &MyVariant);
            if (error<0) goto Error;
            error = CA_VariantGetDispatch (&MyVariant, &MyDispatch);
            if (error<0) goto Error;
            
         
            error = CA_CreateObjHandleFromIDispatch (MyDispatch, 0, &ExcelSingleCellRangeHandle);
            if (error<0) goto Error;
            
        
            error = Excel_GetProperty (ExcelSingleCellRangeHandle, &ErrorInfo, Excel_RangeValue2, CAVT_VARIANT, &MyVariant);
            if (error<0) goto Error;
			
		    if(j == wifi_link_qual_col)
			{
			   test_Type=1; 
			   
               if(CA_VariantHasDouble(&MyVariant)==1)
			   {
	            	error = CA_VariantGetDouble(&MyVariant, &int_numb);
					if (error<0) goto Error;
				
					CA_VariantClear(&MyVariant);
		       
					buffer_line[0][4]= (int)int_numb ; 
			   
			   }
               else if(CA_VariantHasCString(&MyVariant))
			   {
	            	error = CA_VariantGetCString(&MyVariant, &str_status);
					if (error<0) goto Error;
				
					CA_VariantClear(&MyVariant);
		      
			   	    buffer_line[0][4]= 0; 
			   }
			   
			}
			else if(j == wifi_ping_col)
			{
			   test_Type=2; 
			   
               if(CA_VariantHasDouble(&MyVariant)==1)
			   {
	            	error = CA_VariantGetDouble(&MyVariant, &int_numb);
					if (error<0) goto Error;
				
					CA_VariantClear(&MyVariant);
		       
					buffer_line[1][4]= (int)int_numb ; 
			   
			   }
               else if(CA_VariantHasCString(&MyVariant))
			   {
	            	error = CA_VariantGetCString(&MyVariant, &str_status);
					if (error<0) goto Error;
				
					CA_VariantClear(&MyVariant);
		      
			   	   buffer_line[1][4]= -1; 
			   }
			
			}
			
			else if(j == wifi_signallevel_col)
			{
			   test_Type=3;
				
               if(CA_VariantHasDouble(&MyVariant)==1)
			   {
	            	error = CA_VariantGetDouble(&MyVariant, &int_numb);
					if (error<0) goto Error;
				
					CA_VariantClear(&MyVariant);
		       
					buffer_line[2][4]= (int)int_numb ; 
			   
			   }
               else if(CA_VariantHasCString(&MyVariant))
			   {
	            	error = CA_VariantGetCString(&MyVariant, &str_status);
					if (error<0) goto Error;
				
					CA_VariantClear(&MyVariant);
		      
			   	   buffer_line[2][4]= 0; 
			   }
			
			}
			else if(test_Type ==1)
			{
				if(j == status_col[0])
				{
			  	   buffer_line[0][3]=1;   //  was buffer_line[0][3]=1;
				   
	               if(CA_VariantHasCString(&MyVariant))
				   {
		            	error = CA_VariantGetCString(&MyVariant, &str_status);
						if (error<0) goto Error;
				
						CA_VariantClear(&MyVariant);
		    
						if(strcmp(str_status,"PASS")==0)
						{
						   buffer_line[0][2]= 10;
						   number_pass++;
						}
						else if(strcmp(str_status,"FAIL")==0)
						{
						   buffer_line[0][2]= 12;
						   number_fail++;
						}
						else if(strcmp(str_status,"PART_FAIL")==0)
						{
						   buffer_line[0][2]= 13;
						   number_Part_fail++;
						}
						else
						{
						
					
						}
			   
				   }
			    
				}
				else if(j == group_col[0])
				{
					   buffer_line[0][3]=1;
					   
		               if(CA_VariantHasDouble(&MyVariant)==1)
					   {
			            	error = CA_VariantGetDouble(&MyVariant, &int_numb);
							if (error<0) goto Error;
				
							CA_VariantClear(&MyVariant);
		
							buffer_line[0][0]= (int)int_numb;  
					   }
				}
				else if(j == num_col[0])
				{
					   buffer_line[0][3]=1;
					   
		               if(CA_VariantHasDouble(&MyVariant)==1)
					   {
			            	error = CA_VariantGetDouble(&MyVariant, &int_numb);
							if (error<0) goto Error;
				
							CA_VariantClear(&MyVariant);
							
					   		buffer_line[0][1]= (int)int_numb; 
					   }
					   
					   test_Type ==0; 
				}
			}
			else if(test_Type ==2)
			{
				if(j == status_col[1])
				{
			  	   buffer_line[1][3]=2;
				   
	               if(CA_VariantHasCString(&MyVariant))
				   {
		            	error = CA_VariantGetCString(&MyVariant, &str_status);
						if (error<0) goto Error;
				
						CA_VariantClear(&MyVariant);
		    
						if(strcmp(str_status,"PASS")==0)
						{
						   buffer_line[1][2]= 10;
						   number_pass++;
						}
						else if(strcmp(str_status,"FAIL")==0)
						{
						   buffer_line[1][2]= 12;
						   number_fail++;
						}
						else if(strcmp(str_status,"PART_FAIL")==0)
						{
						   buffer_line[1][2]= 13;
						   number_Part_fail++;
						}
			   
				   }
			    
				}
				else if(j == group_col[1])
				{
					   buffer_line[1][3]=2;
					
		               if(CA_VariantHasDouble(&MyVariant)==1)
					   {
			            	error = CA_VariantGetDouble(&MyVariant, &int_numb);
							if (error<0) goto Error;
				
							CA_VariantClear(&MyVariant);
		
							buffer_line[1][0]= (int)int_numb;  
					   }
				}
				else if(j == num_col[1])
				{
					   buffer_line[1][3]=2;
					   
		               if(CA_VariantHasDouble(&MyVariant)==1)
					   {
			            	error = CA_VariantGetDouble(&MyVariant, &int_numb);
							if (error<0) goto Error;
				
							CA_VariantClear(&MyVariant);
							
					   		buffer_line[1][1]= (int)int_numb; 
					   }
					   
					   test_Type ==0; 
				}
			 }
			 else if(test_Type ==3)
			 {
				if(j == status_col[2])
				{
			  	   buffer_line[2][3]=3;
					   
	               if(CA_VariantHasCString(&MyVariant))
				   {
		            	error = CA_VariantGetCString(&MyVariant, &str_status);
						if (error<0) goto Error;
				
						CA_VariantClear(&MyVariant);
		    
						if(strcmp(str_status,"PASS")==0)
						{
						   buffer_line[2][2]= 10;
						   number_pass++;
						}
						else if(strcmp(str_status,"FAIL")==0)
						{
						   buffer_line[2][2]= 12;
						   number_fail++;
						}
						else if(strcmp(str_status,"PART_FAIL")==0)
						{
						   buffer_line[2][2]= 13;
						   number_Part_fail++;
						}
				   }
			    
				}
				else if(j == group_col[2])
				{
					buffer_line[2][3]=3;
					
	               if(CA_VariantHasDouble(&MyVariant)==1)
				   {
		            	error = CA_VariantGetDouble(&MyVariant, &int_numb);
						if (error<0) goto Error;
			
						CA_VariantClear(&MyVariant);
	
						buffer_line[2][0]= (int)int_numb;  
				   }
			    }
				else if(j == num_col[2])
				{
				   buffer_line[2][3]=3;
				   
	               if(CA_VariantHasDouble(&MyVariant)==1)
				   {
		            	error = CA_VariantGetDouble(&MyVariant, &int_numb);
						if (error<0) goto Error;
			
						CA_VariantClear(&MyVariant);
						
				   		buffer_line[2][1]= (int)int_numb; 
				   }
				   
				   test_Type ==0; 
				}
			 }
				
		   }
			
	   }
		 
	   if(flag_process_data == 1)
	   {
	    ret = process_Value_row(buffer_line ,MAX_ROW,MAX_COL, buf_Pod);
		
	    flag_process_data = 0; 
	   }
	}
	
	copy_Pod(buf_Pod,auto_Pod,MAX_ROW,MAX_COL,MAX_POD);
	
	Error:
    SetWaitCursor (0);
    
    CA_VariantClear(&MyVariant);
    CA_VariantClear(&MyCellRangeV);
    
    // Free array of VARAINT
    if (vArray)
        CA_FreeMemory(vArray);
        
    // Free Range handles
    ClearObjHandle (&ExcelRangeHandle);
    ClearObjHandle (&ExcelSingleCellRangeHandle);
    
    if (error < 0) 
        ReportAppAutomationError (error);
	
	
	// --- free allocation pod_buff --
	
	for(i=0 ; i<MAX_POD; i++)
	{
	
	  for(j=0 ; j<MAX_ROW; j++)
	  {
	    free(buf_Pod[i][j]);
	  }
	
	}
	
	free(buf_Pod);
	
    return error;
}

//============================================================================================
//
//============================================================================================

void get_Value_row_col(int *row , int *col)
{
   int ret;
   int rowValue;
   int colValue;
   
   ret = GetCtrlVal(panelHandle,PANEL_NUMERIC_ROW,&rowValue);
   ret = GetCtrlVal(panelHandle,PANEL_NUMERIC_COLUMN,&colValue); 
   *row=rowValue;
   *col=colValue; 

}

//=============================================================================================
//
//=============================================================================================

void copy_Pod(int **auto_Pod[MAX_POD] ,int **auto_Pod_Out[MAX_POD], int row, int col, int num_Pod)
{
   int i;
   int j;
   int k;
   
   for(i=0 ; i<num_Pod; i++)
	{

	   for(j=0 ; j<row; j++)
	   {
	   
		  for(k=0 ; k<col; k++)
		  {
		   auto_Pod_Out[i][j][k]=auto_Pod[i][j][k];
		  }
	   
	   }
	
	}
}

//=============================================================================================
//
//=============================================================================================

int process_Value_row(int buf[MAX_TYPE][BUF_LEN],int row , int col, int **outPOD[MAX_POD])	 // int process_Value_row(int buf[BUF_LEN],int outPOD[MAX_POD][MAX_ROW][MAX_COL])	
{
   int ret=0;
   
   int i =0;
	   
   int j =0; 
   
   int k =0; 
	   ;
   for(i=0 ; i < MAX_TYPE ; i++)
   {
   
      switch (buf[i][0])
      {
	     case POD1:
		 
				 if(buf[i][1] == JIG1)
				 {
					   outPOD[0][0][0] +=1;
		   
					  if(i == 0)
					  {
						   outPOD[0][0][4] =buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[0][0][1] +=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI) 
						   {
						     outPOD[0][0][2]+=1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[0][0][3]+=1;  
						   }
					  }
					  else if(i == 1)
					  {
						   outPOD[0][0][8] =buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[0][0][5] +=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI) 
						   {
						     outPOD[0][0][6]+=1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[0][0][7]+=1;  
						   }
					  
					  }
					  else if( i == 2)
					  {
                           outPOD[0][0][12] =buf[i][3];
							
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[0][0][9] +=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI) 
						   {
						     outPOD[0][0][10]+=1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[0][0][11]+=1;  
						   }
					  
					  }
				 }
				 else if(buf[i][1] == JIG2)
				 {
					   
				   	   outPOD[0][1][0]+=1;  
					   
					   if(i == 0)
					   {
						   outPOD[0][1][4]=buf[i][3];  
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[0][1][1]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[0][1][2]+=1;
						   }
						   else if(buf[i][4]==PART_FAIL_WIFI)
						   {
						     outPOD[0][1][3]+=1;
						   }
					   
					   }
					   else if(i == 1)
					   {
						   outPOD[0][1][8]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[0][1][5]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[0][1][6]+=1;
						   }
						   else if(buf[i][4]==PART_FAIL_WIFI)
						   {
						     outPOD[0][1][7]+=1;
						   }
					   
					   }
					   else if(i == 2)
					   {
						   outPOD[0][1][12]=buf[i][3];  
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[0][1][9]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[0][1][10]+=1;
						   }
						   else if(buf[i][4]==PART_FAIL_WIFI)
						   {
						     outPOD[0][1][11]+=1;
						   }
					   
					   }
		   
				   
				 }
				 else if(buf[i][1] == JIG3)
				 {
					 outPOD[0][2][0]+=1;
					 
					   if(i == 0)
					   {
						   outPOD[0][2][4]=buf[i][3];
				   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[0][2][1]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[0][2][2]+=1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI) 
						   {
						     outPOD[0][2][3]+=1;
						   }
						   
					   }
					   else if(i == 1)
					   {
						   outPOD[0][2][8]=buf[i][3];
				   
				   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[0][2][1]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[0][2][2]+=1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI) 
						   {
						     outPOD[0][2][3]+=1;
						   }
					     
					   }
					   else if(i == 2)
					   {
						   outPOD[0][2][12]=buf[i][3];
				   
				   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[0][2][1]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[0][2][2]+=1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI) 
						   {
						     outPOD[0][2][3]+=1;
						   }
					   
					   }
		  
				 }
				 else if(buf[i][1] == JIG4)
				 {
					   outPOD[0][3][0]+=1;
				   
					   if( i == 0)
					   {
		   				   outPOD[0][3][4]=buf[i][3];
		   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[0][3][1]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI) 
						   {
						     outPOD[0][3][2]+=1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[0][3][3]+=1;
						   }
						 
					   }
					   else if(i == 1)
					   {
		   				 outPOD[0][3][8]=buf[i][3];
		   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[0][3][5]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI) 
						   {
						     outPOD[0][3][6]+=1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[0][3][7]+=1;
						   }
						 
					   }
					   else if(i == 2)
					   {
		   				 outPOD[0][3][12]=buf[i][3];
		   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[0][3][9]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI) 
						   {
						     outPOD[0][3][10]+=1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[0][3][12]+=1;
						   }
						 
					   }
					   
				 }
		 
		 break;
	     case POD2:
		 
				 if(buf[i][1] == JIG1)
				 {  
					   outPOD[1][0][0] +=1;
					   
					   if(i== 0)
					   {
					       outPOD[1][0][4]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[1][0][1]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[1][0][2]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[1][0][3]=+1;
						   }
						 
					   }
					   else if(i==1)
					   {
					       outPOD[1][0][8]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[1][0][5]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[1][0][6]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[1][0][7]=+1;
						   }
						 
					   }
					   else if(i==2)
					   {
					       outPOD[1][0][12]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[1][0][9]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[1][0][10]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[1][0][11]=+1;
						   }
						 
					   
					   }
				   
			
				 }
				 else if(buf[i][1] == JIG2)
				 {
					   outPOD[1][1][0]+=1; 
					   
					   if(i== 0)
					   {
					        outPOD[1][1][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[1][1][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[1][1][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[1][1][3]+=1;	
							}
						 
					   }
					   else if(i==1)
					   {
					        outPOD[1][1][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[1][1][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[1][1][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[1][1][7]+=1;	
							}
						 
					   }
					   else if(i==2)
					   {
					        outPOD[1][1][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[1][1][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[1][1][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[1][1][11]+=1;	
							}
					   
					   }
			
				 }
				 else if(buf[i][1] == JIG3)
				 {
					   outPOD[1][2][0]+=1;
					
					   if(i== 0)
					   {
					     	outPOD[1][2][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[1][2][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[1][2][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[1][2][3]+=1;
							}
							
					   }
					   else if(i==1)
					   {
					   	   	outPOD[1][2][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[1][2][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[1][2][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[1][2][7]+=1;
							}
							
					   }
					   else if(i==2)
					   {
					   	    outPOD[1][2][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[1][2][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[1][2][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[1][2][11]+=1;
							}
					   
					   }
						
				 }
				 else if(buf[i][1] == JIG4)
				 {
					    outPOD[1][3][0]+=1;
					
					   if(i== 0)
					   {
					        outPOD[1][3][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[1][3][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[1][3][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							   outPOD[1][3][3]+=1;
							}
						  
					   }
					   else if(i==1)
					   {
					        outPOD[1][3][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[1][3][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[1][3][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							   outPOD[1][3][7]+=1;
							}
						  
					   }
					   else if(i==2)
					   {
					        outPOD[1][3][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[1][3][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[1][3][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							   outPOD[1][3][11]+=1;
							}
					   
					   }
						
				 }
		 
		 break;
	     case POD3:
		 
				 if(buf[i][1] == JIG1)
				 {  
					   outPOD[2][0][0] +=1;
					   
					   if(i== 0)
					   {
					       outPOD[2][0][4]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[2][0][1]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[2][0][2]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[2][0][3]=+1;
						   }
						 
					   }
					   else if(i==1)
					   {
					       outPOD[2][0][8]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[2][0][5]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[2][0][6]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[2][0][7]=+1;
						   }
						 
					   }
					   else if(i==2)
					   {
					       outPOD[2][0][12]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[2][0][9]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[2][0][10]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[2][0][11]=+1;
						   }
						 
					   
					   }
				   
			
				 }
				 else if(buf[i][1] == JIG2)
				 {
					   outPOD[2][1][0]+=1; 
					   
					   if(i== 0)
					   {
					        outPOD[2][1][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[2][1][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[2][1][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[2][1][3]+=1;	
							}
						 
					   }
					   else if(i==1)
					   {
					        outPOD[2][1][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[2][1][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[2][1][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[2][1][7]+=1;	
							}
						 
					   }
					   else if(i==2)
					   {
					        outPOD[2][1][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[2][1][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[2][1][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[2][1][11]+=1;	
							}
					   
					   }
			
				 }
				 else if(buf[i][1] == JIG3)
				 {
					   outPOD[2][2][0]+=1;
					
					   if(i== 0)
					   {
					     	outPOD[2][2][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[2][2][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[2][2][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[2][2][3]+=1;
							}
							
					   }
					   else if(i==1)
					   {
					   	   	outPOD[2][2][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[2][2][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[2][2][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[2][2][7]+=1;
							}
							
					   }
					   else if(i==2)
					   {
					   	    outPOD[2][2][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[2][2][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[2][2][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[2][2][11]+=1;
							}
					   
					   }
						
				 }
				 else if(buf[i][1] == JIG4)
				 {
					    outPOD[2][3][0]+=1;
					
					   if(i== 0)
					   {
					        outPOD[2][3][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[2][3][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[2][3][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							   outPOD[2][3][3]+=1;
							}
						  
					   }
					   else if(i==1)
					   {
					        outPOD[2][3][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[2][3][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[2][3][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							   outPOD[2][3][7]+=1;
							}
						  
					   }
					   else if(i==2)
					   {
					        outPOD[2][3][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[2][3][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[2][3][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							   outPOD[2][3][11]+=1;
							}
					   
					   }
						
				 }
		 	     
		 break;
	 
	     case POD4:
				 if(buf[i][1] == JIG1)
				 {  
					   outPOD[3][0][0] +=1;
					   
					   if(i== 0)
					   {
					       outPOD[3][0][4]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[3][0][1]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[3][0][2]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[3][0][3]=+1;
						   }
						 
					   }
					   else if(i==1)
					   {
					       outPOD[3][0][8]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[3][0][5]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[3][0][6]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[3][0][7]=+1;
						   }
						 
					   }
					   else if(i==2)
					   {
					       outPOD[3][0][12]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[3][0][9]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[3][0][10]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[3][0][11]=+1;
						   }
						 
					   
					   }
				   
			
				 }
				 else if(buf[i][1] == JIG2)
				 {
					   outPOD[3][1][0]+=1; 
					   
					   if(i== 0)
					   {
					        outPOD[3][1][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[3][1][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[3][1][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[3][1][3]+=1;	
							}
						 
					   }
					   else if(i==1)
					   {
					        outPOD[3][1][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[3][1][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[3][1][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[3][1][7]+=1;	
							}
						 
					   }
					   else if(i==2)
					   {
					        outPOD[3][1][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[3][1][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[3][1][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[3][1][11]+=1;	
							}
					   
					   }
			
				 }
				 else if(buf[i][1] == JIG3)
				 {
					   outPOD[3][2][0]+=1;
					
					   if(i== 0)
					   {
					     	outPOD[3][2][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[3][2][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[3][2][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[3][2][3]+=1;
							}
							
					   }
					   else if(i==1)
					   {
					   	   	outPOD[3][2][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[3][2][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[3][2][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[3][2][7]+=1;
							}
							
					   }
					   else if(i==2)
					   {
					   	    outPOD[3][2][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[3][2][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[3][2][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[3][2][11]+=1;
							}
					   
					   }
						
				 }
				 else if(buf[i][1] == JIG4)
				 {
					    outPOD[3][3][0]+=1;
					
					   if(i== 0)
					   {
					        outPOD[3][3][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[3][3][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[3][3][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							   outPOD[3][3][3]+=1;
							}
						  
					   }
					   else if(i==1)
					   {
					        outPOD[3][3][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[3][3][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[3][3][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							   outPOD[3][3][7]+=1;
							}
						  
					   }
					   else if(i==2)
					   {
					        outPOD[3][3][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[3][3][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[3][3][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							   outPOD[3][3][11]+=1;
							}
					   
					   }
						
				 }
		 
		 
		 break;
	 
	     case POD5:
				 if(buf[i][1] == JIG1)
				 {  
					   outPOD[4][0][0] +=1;
					   
					   if(i== 0)
					   {
					       outPOD[4][0][4]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[4][0][1]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[4][0][2]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[4][0][3]=+1;
						   }
						 
					   }
					   else if(i==1)
					   {
					       outPOD[4][0][8]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[4][0][5]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[4][0][6]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[4][0][7]=+1;
						   }
						 
					   }
					   else if(i==2)
					   {
					       outPOD[4][0][12]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[4][0][9]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[4][0][10]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[4][0][11]=+1;
						   }
						 
					   
					   }
				   
			
				 }
				 else if(buf[i][1] == JIG2)
				 {
					   outPOD[4][1][0]+=1; 
					   
					   if(i== 0)
					   {
					        outPOD[4][1][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[4][1][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[4][1][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[4][1][3]+=1;	
							}
						 
					   }
					   else if(i==1)
					   {
					        outPOD[4][1][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[4][1][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[4][1][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[4][1][7]+=1;	
							}
						 
					   }
					   else if(i==2)
					   {
					        outPOD[4][1][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[4][1][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[4][1][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[4][1][11]+=1;	
							}
					   
					   }
			
				 }
				 else if(buf[i][1] == JIG3)
				 {
					   outPOD[4][2][0]+=1;
					
					   if(i== 0)
					   {
					     	outPOD[4][2][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[4][2][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[4][2][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[4][2][3]+=1;
							}
							
					   }
					   else if(i==1)
					   {
					   	   	outPOD[4][2][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[4][2][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[4][2][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[4][2][7]+=1;
							}
							
					   }
					   else if(i==2)
					   {
					   	    outPOD[4][2][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[4][2][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[4][2][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[4][2][11]+=1;
							}
					   
					   }
						
				 }
				 else if(buf[i][1] == JIG4)
				 {
					    outPOD[4][3][0]+=1;
					
					   if(i== 0)
					   {
					        outPOD[4][3][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[4][3][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[4][3][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							   outPOD[4][3][3]+=1;
							}
						  
					   }
					   else if(i==1)
					   {
					        outPOD[4][3][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[4][3][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[4][3][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							   outPOD[4][3][7]+=1;
							}
						  
					   }
					   else if(i==2)
					   {
					        outPOD[4][3][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[4][3][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI)
						    {
						      outPOD[4][3][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							   outPOD[4][3][11]+=1;
							}
					   
					   }
						
				 }
		 	   
		 break;
	 
	     case MANUALJIG_1 :
		 
				 if(buf[i][1] == JIG1)
				 {  
					   outPOD[5][0][0] +=1;
					   
					   if(i== 0)
					   {
					       outPOD[5][0][4]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[5][0][1]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[5][0][2]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[5][0][3]=+1;
						   }
						 
					   }
					   else if(i==1)
					   {
					       outPOD[5][0][8]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[5][0][5]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[5][0][6]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[5][0][7]=+1;
						   }
						 
					   }
					   else if(i==2)
					   {
					       outPOD[5][0][12]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[5][0][9]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[5][0][10]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[5][0][11]=+1;
						   }
						 
					   
					   }
				   
			
				 }
				 else if(buf[i][1] == JIG2)
				 {
					   outPOD[5][1][0]+=1; 
					   
					   if(i== 0)
					   {
					        outPOD[5][1][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[5][1][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[5][1][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[5][1][3]+=1;	
							}
						 
					   }
					   else if(i==1)
					   {
					        outPOD[5][1][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[5][1][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[5][1][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[5][1][7]+=1;	
							}
						 
					   }
					   else if(i==2)
					   {
					        outPOD[5][1][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[5][1][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[5][1][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[5][1][11]+=1;	
							}
					   
					   }
			
				 }
		 break;
	 
	     case MANUALJIG_2 :
			 
				 if(buf[i][1] == JIG1)
				 {  
					   outPOD[6][0][0] +=1;
					   
					   if(i== 0)
					   {
					       outPOD[6][0][4]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[6][0][1]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[6][0][2]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[6][0][3]=+1;
						   }
						 
					   }
					   else if(i==1)
					   {
					       outPOD[6][0][8]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[6][0][5]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[6][0][6]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[6][0][7]=+1;
						   }
						 
					   }
					   else if(i==2)
					   {
					       outPOD[6][0][12]=buf[i][3];
						   
						   if(buf[i][2]==PASS_WIFI)
						   {
						     outPOD[6][0][9]+=1; 
						   }
						   else if(buf[i][2]==FAIL_WIFI)
						   {
						     outPOD[6][0][10]=+1;
						   }
						   else if(buf[i][2]== PART_FAIL_WIFI)
						   {
						     outPOD[6][0][11]=+1;
						   }
						 
					   
					   }
				   
			
				 }
				 else if(buf[i][1] == JIG2)
				 {
					   outPOD[6][1][0]+=1; 
					   
					   if(i== 0)
					   {
					        outPOD[6][1][4]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[6][1][1]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[6][1][2]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[6][1][3]+=1;	
							}
						 
					   }
					   else if(i==1)
					   {
					        outPOD[6][1][8]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[6][1][5]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[6][1][6]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[6][1][7]+=1;	
							}
						 
					   }
					   else if(i==2)
					   {
					        outPOD[6][1][12]=buf[i][3];
							
						    if(buf[i][2]==PASS_WIFI)
						    {
						      outPOD[6][1][9]+=1; 
						    }
						    else if(buf[i][2]==FAIL_WIFI) 
						    {
						      outPOD[6][1][10]+=1;
						    }
							else if(buf[i][2]== PART_FAIL_WIFI)
							{
							  outPOD[6][1][11]+=1;	
							}
					   
					   }
			
				 }
		       
		 break;
		 default:
			printf("default");
   
	   }
   
   
   }
   
   return 0; 
}


//=============================================================================================
//
//=============================================================================================

int display_Result_Pi(int **auto_Pod[MAX_POD], int row , int col , int num_Pod)
{
  int i;
  int j;
  int k;
  char str_Num_Jig[15];
  char str_Text_Jig[100];
  char buf_Jig_Pi[250];
  for(i=0 ; i<MAX_POD ; i++)
  {
     InsertTextBoxLine(panelHandle,PANEL_TEXTBOX_POD,-1,autopod_Pi[i]); 
	 SetCtrlAttribute (panelHandle, PANEL_TEXTBOX_POD, ATTR_TEXT_BOLD, 0);
	 SetCtrlAttribute (panelHandle, PANEL_TEXTBOX_POD,ATTR_TEXT_POINT_SIZE, 14);
	 InsertTextBoxLine(panelHandle,PANEL_TEXTBOX_POD,-1,"");
	 
	 for(j=0 ; j<MAX_JIG ; j++)
	 {
		   sprintf(str_Num_Jig,"%d",auto_Pod[i][j][0]);
		   strcpy(str_Text_Jig," Boards.");
		   strcpy(buf_Jig_Pi,jig_pi[j]); 
		   InsertTextBoxLine(panelHandle,PANEL_TEXTBOX_POD,-1,strcat(buf_Jig_Pi,strcat(str_Num_Jig,str_Text_Jig)) );
		   InsertTextBoxLine(panelHandle,PANEL_TEXTBOX_POD,-1,"");
	   
		   for(k=0 ; k<MAX_TYPE ; k++)
		   {
	   
			   if(auto_Pod[i][j][4] == k+1)                    // k+1 = 1 = WIFI_QUALITY 
			   {
					char str_Pass[80];
			        char str_Fail[30];
					char str_Part_Fail[30];
			        char str_tab_Link_Quality[650]={"                      "}; 
					sprintf(str_Pass,"%d",auto_Pod[i][j][1]); 
				    sprintf(str_Fail,"%d",auto_Pod[i][j][2]);
					sprintf(str_Part_Fail,"%d",auto_Pod[i][j][3]);
					strcat(str_Pass,"  PASSED             ");
				    strcat(str_Fail,"  FAILED    ");
					strcat(str_Part_Fail,"   PART_FAILED .");
				    strcat(str_tab_Link_Quality,wifi_Type_Test[k]);                      // wifi_Type_Test[0] = WIFI_LINL_QUALITY 
				    InsertTextBoxLine(panelHandle,PANEL_TEXTBOX_POD,-1,strcat(str_tab_Link_Quality,strcat(strcat(str_Pass,str_Fail),str_Part_Fail)));
					InsertTextBoxLine(panelHandle,PANEL_TEXTBOX_POD,-1,"");
			   }
			   else if(auto_Pod[i][j][8] == k+1)             //   k+1 = 2 = TEST WIFI_PINF
			   {
	   
			        char str_Pass[60];
			        char str_Fail[30];
					char str_Part_Fail[30];
			        char str_tab_Wifi_Ping[650]   ={"                      "};   // shif by 8
					sprintf(str_Pass,"%d",auto_Pod[i][j][5]); 
				    sprintf(str_Fail,"%d",auto_Pod[i][j][6]);    
					sprintf(str_Part_Fail,"%d",auto_Pod[i][j][7]);
					strcat(str_Pass,"  PASSED             ");
				    strcat(str_Fail,"  FAILED    ");
					strcat(str_Part_Fail,"   PART_FAILED .");
				    strcat(str_tab_Wifi_Ping,wifi_Type_Test[k]);                       
				    //InsertTextBoxLine(panelHandle,PANEL_TEXTBOX_POD,-1,strcat(str_tab_Wifi_Ping,strcat(str_Pass,str_Fail)));
					InsertTextBoxLine(panelHandle,PANEL_TEXTBOX_POD,-1,strcat(str_tab_Wifi_Ping,strcat(strcat(str_Pass,str_Fail),str_Part_Fail))); 
					InsertTextBoxLine(panelHandle,PANEL_TEXTBOX_POD,-1,"");
	   
	   
			   }
			   else if(auto_Pod[i][j][12] == k+1 )			 //    k+1 = 3 = WIFI_SIGNSL_LEVEL
			   {
			        char str_Pass[60];
			        char str_Fail[30];
					char str_Part_Fail[30];
			        char str_tab_Signal_Level[650]={"                      "}; 
					sprintf(str_Pass,"%d",auto_Pod[i][j][9]); 
				    sprintf(str_Fail,"%d",auto_Pod[i][j][10]);    
					sprintf(str_Part_Fail,"%d",auto_Pod[i][j][11]);
					strcat(str_Pass,"  PASSED             ");
				    strcat(str_Fail,"  FAILED    ");
					strcat(str_Part_Fail,"   PART_FAILED .");
				    strcat(str_tab_Signal_Level,wifi_Type_Test[k]);                      
				    //InsertTextBoxLine(panelHandle,PANEL_TEXTBOX_POD,-1,strcat(str_tab_Signal_Level,strcat(str_Pass,str_Fail)));
					InsertTextBoxLine(panelHandle,PANEL_TEXTBOX_POD,-1,strcat(str_tab_Signal_Level,strcat(strcat(str_Pass,str_Fail),str_Part_Fail))); 
					InsertTextBoxLine(panelHandle,PANEL_TEXTBOX_POD,-1,"");
	   
			   }
	  
	       }
	   
	  }
  
  }
  
  

  return 0;
}
//====================================================================================
//
//====================================================================================
int init_Globla_var(void)
{

	 int i;
	 int j;
	 int k;
	 
	 auto_Pod_Pi = (int***)malloc(sizeof(int**)*MAX_POD);
	
	if(auto_Pod_Pi == NULL)
	{
	  return -1; 
	}
	
	for(i=0 ; i<MAX_POD; i++)
	{
	   auto_Pod_Pi[i]=(int**)malloc(sizeof(int*)*MAX_ROW);
	   
	   if(auto_Pod_Pi[i]==NULL)
	   {
	     return -1;
	   }
	   
	   for(j=0 ; j<MAX_ROW; j++)
	   {
	      auto_Pod_Pi[i][j]= (int*)malloc(sizeof(int)*MAX_COL);
		  
		  if(auto_Pod_Pi[i][j]==NULL)
		  {
		    return -1; 
		  }
		  
		  for(k=0 ; k<MAX_COL; k++)
		  {
		    auto_Pod_Pi[i][j][k]=0;
		  }
	   
	   }
	
	}
   return 0;
}

//====================================================================================
//
//====================================================================================
void free_Global_Var(void)
{
	int i;
	int j;
 	for(i=0 ; i<MAX_POD; i++)
	{
	
	  for(j=0 ; j<MAX_ROW; j++)
	  {
	    free(auto_Pod_Pi[i][j]);
	  }
	
	}
	
	free(auto_Pod_Pi);
}
