// DDEClient.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"

#include "windows.h"
#include "ddeml.h"
#include "stdio.h"

HDDEDATA CALLBACK DdeCallback(
    UINT uType,     // Transaction type.
    UINT uFmt,      // Clipboard data format.
    HCONV hconv,    // Handle to the conversation.
    HSZ hsz1,       // Handle to a string.
    HSZ hsz2,       // Handle to a string.
    HDDEDATA hdata, // Handle to a global memory object.
    DWORD dwData1,  // Transaction-specific data.
    DWORD dwData2)  // Transaction-specific data.
{
    return 0;
}

void DDEExecute(DWORD idInst, HCONV hConv, char* szCommand)
{
    HDDEDATA hData = DdeCreateDataHandle(idInst, (LPBYTE)szCommand,
                               lstrlen(szCommand)+1, 0, NULL, CF_TEXT, 0);
    if (hData==NULL)   {
        printf("Command failed: %s\n", szCommand);
    }
    else    {
        DdeClientTransaction((LPBYTE)hData, 0xFFFFFFFF, hConv, 0L, 0,
                             XTYP_EXECUTE, TIMEOUT_ASYNC, NULL);
    }
}

void DDERequest(DWORD idInst, HCONV hConv, char* szItem, char* sDesc)
{
    HSZ hszItem = DdeCreateStringHandle(idInst, szItem, 0);
    HDDEDATA hData = DdeClientTransaction(NULL,0,hConv,hszItem,CF_TEXT, 
                                 XTYP_REQUEST,5000 , NULL);
    if (hData==NULL)
    {
        printf("Request failed: %s\n", szItem);
    }
    else
    {
        char szResult[255];
        DdeGetData(hData, (unsigned char *)szResult, 255, 0);
        printf("%s%s\n", sDesc, szResult);
    }
}

void DDEPoke(DWORD idInst, HCONV hConv, char* szItem, char* szData)
{
    HSZ hszItem = DdeCreateStringHandle(idInst, szItem, 0);
	DdeClientTransaction((LPBYTE)szData, (DWORD)(lstrlen(szData)+1),
                          hConv, hszItem, CF_TEXT,
                          XTYP_POKE, 3000, NULL);
    DdeFreeStringHandle(idInst, hszItem);
}

int _tmain(int argc, _TCHAR* argv[])
{
	char szApp[] = "EXCEL";
    char szTopic[] = "C:\\Test.xls";
    char szCmd1[] = "[APP.MINIMIZE()]";
    char szItem1[] = "R1C1";  char szDesc1[] = "A1 Contains: ";
    char szItem2[] = "R2C1";  char szDesc2[] = "A2 Contains: ";
    char szItem3[] = "R3C1";  char szData3[] = "Data from DDE Client";
    char szCmd2[] = "[SELECT(\"R3C1\")][FONT.PROPERTIES(,\"Bold\")][SAVE()][QUIT()]";

    //DDE Initialization
    DWORD idInst=0;
    UINT iReturn;
    iReturn = DdeInitialize(&idInst, (PFNCALLBACK)DdeCallback, 
                            APPCLASS_STANDARD | APPCMD_CLIENTONLY, 0 );
    if (iReturn!=DMLERR_NO_ERROR)
    {
        printf("DDE Initialization Failed: 0x%04x\n", iReturn);
        Sleep(1500);
        return 0;
    }

    //Start DDE Server and wait for it to become idle.
    //HINSTANCE hRet = ShellExecute(0, "open", szTopic, 0, 0, SW_SHOWNORMAL);
    //if ((int)hRet < 33)
    //{
    //    printf("Unable to Start DDE Server: 0x%04x\n", hRet);
    //    Sleep(1500); DdeUninitialize(idInst);
    //    return 0;
    //}
    //Sleep(1000);

    //DDE Connect to Server using given AppName and topic.
    HSZ hszApp, hszTopic;
    HCONV hConv;
    hszApp = DdeCreateStringHandle(idInst, szApp, 0);
    hszTopic = DdeCreateStringHandle(idInst, szTopic, 0);
    hConv = DdeConnect(idInst, hszApp, hszTopic, NULL);
    DdeFreeStringHandle(idInst, hszApp);
    DdeFreeStringHandle(idInst, hszTopic);
    if (hConv == NULL)
    {
        printf("DDE Connection Failed.\n");
        Sleep(1500); DdeUninitialize(idInst);
        return 0;
    }

    //Execute commands/requests specific to the DDE Server.
    DDEExecute(idInst, hConv, szCmd1);
    DDERequest(idInst, hConv, szItem1, szDesc1); 
    DDERequest(idInst, hConv, szItem2, szDesc2);
    DDEPoke(idInst, hConv, szItem3, szData3);
    //DDEExecute(idInst, hConv, szCmd2);

    //DDE Disconnect and Uninitialize.
    DdeDisconnect(hConv);
    DdeUninitialize(idInst);

    Sleep(3000);
    return 1;
}

