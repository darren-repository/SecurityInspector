// SecurityInspector.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "SecurityInspector.h"

#include <windowsx.h>
#include <commctrl.h>

#include <AclAPI.h>
#include <Sddl.h>
#include <AccCtrl.h>

#include <string>
#include <vector>

using namespace std;

#define MAX_LOADSTRING 100

//#pragma comment( lib, "Comctl32.lib" )


// Global Variables:
HINSTANCE g_hInstance;                                // current instance
HWND g_hSecurityDialog;
WCHAR g_szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR g_szWindowClass[MAX_LOADSTRING];            // the main window class name

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
BOOL                LoadChildDialog(HINSTANCE, HWND, INT, DLGPROC);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    SecurityDialogProc(HWND, UINT, WPARAM, LPARAM);

BOOL ProcessControlMessage(NMHDR* pControlMessageInfo);
BOOL ProcessFolderControlMessage(NMHDR* pControlMessage);
BOOL ProcessListViewMessages(NMHDR* pControlMessage);
HRESULT TreeViewClearChildren(HWND hTreeViewControl, TVITEM* pItem);
HRESULT TreeViewGetItemPath(__in HWND hTreeViewControl, __in TVITEM* pItem, __out wstring* pstrItemPath);
HRESULT TreeViewAddNewPath(__in HWND hTreeViewControl, __in TVITEM* pItem, wchar_t* szPath);
HTREEITEM TreeViewAddItem(HWND hTreeView, wstring* strItem, HTREEITEM hParentItem = NULL);

HRESULT ListViewClear(HWND hListView);
HRESULT ListViewAddItem(HWND hListView, wstring* strItem);
HRESULT ListViewAddItems(HWND hListView, vector<wstring>* pstrItems);

HRESULT SetProperTokenAuthority();


class CDirectoryManagement
{
private:
public:
    static HRESULT EnumerateDrives(__out vector<wstring>* pDriveList);
    static HRESULT EnumerateFolders(__in wchar_t* szPath, __out vector<wstring>* pFolderList);
    static HRESULT EnumerateFiles(__in wchar_t* szPath, __out vector<wstring>* pFileList);
    static HRESULT GetObjectSDDLInformation(__in const wchar_t* szObjectName, __out wstring* pSDDLInformation);
};


typedef struct _FIXED_SIZE_TOKEN_PRIVILEGES {
    DWORD PrivilegeCount;
    LUID_AND_ATTRIBUTES Privileges[64];
} FIXED_SIZE_TOKEN_PRIVILEGES;


class CProcessToken
{
private:
    static HRESULT GetTokenPrivileges(__out FIXED_SIZE_TOKEN_PRIVILEGES* pFixedSizeTokenPrivileges);

public:
    CProcessToken() {}
    ~CProcessToken() {}

    static HRESULT GetTokenHandle(__out PHANDLE pTokenHandle);
    static HRESULT DoesTokenHavePrivilege(LPCTSTR tszPrivilegeName);
    static HRESULT EnableTokenPrivilege(LPCTSTR tszPrivilegeName);

};



VOID LoadDirectoryList(HWND hDialogWindow);


int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: Place code here.

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, g_szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_SECURITYINSPECTOR, g_szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_SECURITYINSPECTOR));

    MSG msg;

    //Adjust Process Token
    HRESULT hCheck = CProcessToken::EnableTokenPrivilege(SE_SECURITY_NAME);
    hCheck = CProcessToken::EnableTokenPrivilege(SE_BACKUP_NAME);

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int) msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SECURITYINSPECTOR));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_SECURITYINSPECTOR);
    wcex.lpszClassName  = g_szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   g_hInstance = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowW(g_szWindowClass, g_szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_CREATE:
        LoadChildDialog(g_hInstance, hWnd, IDD_DIALOG_SECURITY, SecurityDialogProc);
        break;
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Parse the menu selections:
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            BeginPaint(hWnd, &ps);
            // TODO: Add any drawing code that uses hdc here...
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    case WM_SIZE:
        SendMessage(GetWindow(hWnd, GW_CHILD), WM_SIZE, 0, 0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

BOOL LoadChildDialog(HINSTANCE hInstance, HWND hWindow, INT iDialogID, DLGPROC lpDialogProc)
{
    g_hSecurityDialog = CreateDialog(hInstance, MAKEINTRESOURCE(iDialogID), hWindow, lpDialogProc);
    
    //SetParent(hChildWnd, hWindow);
    ShowWindow(g_hSecurityDialog, SW_SHOW);

    return TRUE;
}

//////////////////////////////////////////////////////////////
// Resize the security dialog and all child controls based
// upon the parent window size
//////////////////////////////////////////////////////////////
VOID ResizeSecurityDialog(HWND hSecurityDialog)
{
    HWND hParentWindow;
    RECT rctClientArea = {};
    RECT rctControlArea = {};

    hParentWindow = GetParent(hSecurityDialog);
    GetClientRect(hParentWindow, &rctClientArea);

    SetWindowPos(hSecurityDialog, HWND_TOP, rctClientArea.left, rctClientArea.top, rctClientArea.right, rctClientArea.bottom, 0);

    //Modify the static text control size & position
    HWND hControl = GetDlgItem(hSecurityDialog, IDC_STATIC_SDDLINFO);

    rctControlArea.left = rctClientArea.left + 10;
    rctControlArea.right = rctClientArea.right - 20;
    rctControlArea.top = rctClientArea.top + 10;
    rctControlArea.bottom = 74;

    MoveWindow(hControl, rctControlArea.left, rctControlArea.top, rctControlArea.right, rctControlArea.bottom, TRUE);

    //Modify the tree view control size & position
    hControl = GetDlgItem(hSecurityDialog, IDC_FOLDERS);

    rctControlArea.left = rctClientArea.left + 10;
    rctControlArea.right = (long)(((float)rctClientArea.right - (float)rctClientArea.left) / 2.75f) - 10;
    rctControlArea.top = rctClientArea.top + 96;
    rctControlArea.bottom = rctClientArea.bottom - 106;

    MoveWindow(hControl, rctControlArea.left, rctControlArea.top, rctControlArea.right, rctControlArea.bottom, TRUE);

    //Modify the list view control size & position
    hControl= GetDlgItem(hSecurityDialog, IDC_LIST_FILES);

    rctControlArea.left = (long)(((float)rctClientArea.right - (float)rctClientArea.left) / 2.75f) + 10;
    rctControlArea.right = rctClientArea.right - rctControlArea.left - 10;
    rctControlArea.top = rctClientArea.top + 96;
    rctControlArea.bottom = rctClientArea.bottom - 106;

    MoveWindow(hControl, rctControlArea.left, rctControlArea.top, rctControlArea.right, rctControlArea.bottom, TRUE);
}

HRESULT TreeViewClearChildren(HWND hTreeViewControl, TVITEM* pItem)
{
    HTREEITEM hItemHandle;
    TVITEM tvItem;
    wchar_t wszItemText[256] = {};

    if (pItem == nullptr)
    {
        return E_INVALIDARG;
    }

    memcpy_s(&tvItem, sizeof(TVITEM), pItem, sizeof(TVITEM));
    tvItem.mask |= TVIF_TEXT;
    tvItem.pszText = &wszItemText[0];
    tvItem.cchTextMax = 256;
    TreeView_GetItem(hTreeViewControl, &tvItem);


    hItemHandle = TreeView_GetChild(hTreeViewControl, pItem->hItem);

    while (hItemHandle != NULL)
    {
        TreeView_DeleteItem(hTreeViewControl, hItemHandle);
        hItemHandle = TreeView_GetChild(hTreeViewControl, pItem->hItem);
    }

    return S_OK;
}

HRESULT TreeViewGetItemPath(__in HWND hTreeViewControl, __in TVITEM* pItem, __out wstring* pstrItemPath)
{
    wchar_t wszItemText[256] = {};
    vector<wstring> vstrBuildPath;
    HTREEITEM hTreeHandle;
    TVITEM tvItem;

    if (hTreeViewControl == 0 || pItem == nullptr || pstrItemPath == nullptr)
    {
        return E_INVALIDARG;
    }

    memset(&tvItem, 0, sizeof(TVITEM));

    //Modify the item to get the text from the item
    tvItem.mask |= TVIF_TEXT | TVIF_HANDLE;
    tvItem.pszText = wszItemText;
    tvItem.cchTextMax = 256;
    tvItem.hItem = pItem->hItem;


    //Get the current item's text
    TreeView_GetItem(hTreeViewControl, &tvItem);
    vstrBuildPath.push_back(wszItemText);

    hTreeHandle = TreeView_GetParent(hTreeViewControl, tvItem.hItem);

    //Prepare to receive the information from the first parent
    memset(wszItemText, 0, sizeof(wchar_t) * 256);
    memset(&tvItem, 0, sizeof(TVITEM));

    tvItem.mask = TVIF_TEXT | TVIF_HANDLE;
    tvItem.pszText = wszItemText;
    tvItem.cchTextMax = 256;
    tvItem.hItem = hTreeHandle;

    //Get each parent's text for building the path
    while (hTreeHandle != NULL)
    {
        //Get the current item's text
        TreeView_GetItem(hTreeViewControl, &tvItem);
        vstrBuildPath.push_back(wszItemText);

        hTreeHandle = TreeView_GetParent(hTreeViewControl, tvItem.hItem);

        //Prepare to receive the next bit of information
        memset(wszItemText, 0, sizeof(wchar_t) * 256);
        memset(&tvItem, 0, sizeof(TVITEM));
    
        tvItem.mask = TVIF_TEXT;
        tvItem.pszText = wszItemText;
        tvItem.cchTextMax = 256;
        tvItem.hItem = hTreeHandle;
    }
    
    //Clear the string before creating the path
    pstrItemPath->clear();

    //Build out the path using the strings collected
    for (int iStrIdx = ((int)vstrBuildPath.size() - 1); iStrIdx >= 0; iStrIdx--)
    {
        pstrItemPath->append(vstrBuildPath[iStrIdx]);
        pstrItemPath->append(L"\\");
    }

    return S_OK;
}

HRESULT TreeViewAddNewPath(__in HWND hTreeViewControl, __in TVITEM* pItem, wchar_t* szPath)
{
    HTREEITEM hFolder;
    vector<wstring> vstrFolderList;
    vector<wstring> vstrSubFolderList;
    wstring strPath;

    strPath.append(szPath);
    strPath.append(L"*");

    CDirectoryManagement::EnumerateFolders(const_cast<wchar_t*>(strPath.c_str()), &vstrFolderList);
    
    for (unsigned int iDir = 0; iDir < vstrFolderList.size(); iDir++)
    {
        if ( (wcscmp(vstrFolderList[iDir].data(), TEXT(".")) == 0) ||
             (wcscmp(vstrFolderList[iDir].data(), TEXT("..")) == 0) )
        {
            continue;
        }

        hFolder = TreeViewAddItem(hTreeViewControl, (wstring*)&vstrFolderList[iDir], pItem->hItem);
    
        //Clear variables
        vstrSubFolderList.clear();
        strPath.clear();

        //Create the path to the subdirectory
        strPath.append(szPath);
        strPath.append(vstrFolderList[iDir]);
        strPath.append(L"\\*");
    
        //Get and add a subdirectory list
        CDirectoryManagement::EnumerateFolders(const_cast<wchar_t*>(strPath.c_str()), &vstrSubFolderList);
        if (vstrSubFolderList.size() > 2)
        {
            for (unsigned int iSubItem = 0; iSubItem < vstrSubFolderList.size(); iSubItem++)
            {
                if ( (wcscmp(vstrSubFolderList[iSubItem].data(), TEXT(".")) == 0) ||
                     (wcscmp(vstrSubFolderList[iSubItem].data(), TEXT("..")) == 0) )
                {
                    continue;
                }

                TreeViewAddItem(hTreeViewControl, (wstring*)&vstrSubFolderList[iSubItem], hFolder);
            }
        }
    }

    return S_OK;
}

BOOL ProcessFolderControlMessage(NMHDR* pControlMessage)
{
    LPNMTREEVIEW pTreeViewMessage;
    LPTVITEM pNewItem;
    LPTVITEM pOldItem;
    wstring strItemPath;

    HWND hListView;
    vector<wstring> vstrFiles;

    pTreeViewMessage = reinterpret_cast<LPNMTREEVIEW>(pControlMessage);
    pNewItem = &(pTreeViewMessage->itemNew);
    pOldItem = &(pTreeViewMessage->itemNew);

    switch (pControlMessage->code)
    {
    case TVN_ITEMEXPANDING:
        {
            switch (pTreeViewMessage->action)
            {
            case TVE_COLLAPSE:
            case TVE_COLLAPSERESET:
                OutputDebugStringW(L"TVN_ITEMEXPANDING ... TVE_COLLAPSE\r\n");
                break;
            case TVE_EXPAND:
                //Clear the subtree out
                TreeViewClearChildren(pControlMessage->hwndFrom, pNewItem);
                //Get the expanding nodes path
                TreeViewGetItemPath(pControlMessage->hwndFrom, pNewItem, &strItemPath);
                //Add list of directories and sub-directories
                TreeViewAddNewPath(pControlMessage->hwndFrom, pNewItem, const_cast<wchar_t*>(strItemPath.c_str()));
                OutputDebugStringW(L"TVN_ITEMEXPANDING ... TVE_EXPAND\r\n");
                break;
            case TVE_EXPANDPARTIAL:
                OutputDebugStringW(L"TVN_ITEMEXPANDING ... TVE_EXPANDPARTIAL\r\n");
                break;
            }
        }
        break;
    case TVN_SELCHANGING:
        {
            //Get the expanding nodes path
            TreeViewGetItemPath(pControlMessage->hwndFrom, pNewItem, &strItemPath);
            
            //append wild card '*' for search
            strItemPath.append(L"*");

            //Enumerate the files in the path
            CDirectoryManagement::EnumerateFiles(const_cast<wchar_t*>(strItemPath.data()), &vstrFiles);

            //Remove any items from the list view control
            hListView = GetDlgItem(g_hSecurityDialog, IDC_LIST_FILES);

            if (hListView != 0)
            {
                ListViewClear(hListView);
                ListViewAddItems(hListView, &vstrFiles);
            }
            OutputDebugStringW(L"TVN_SELCHANGING\r\n");

        }
        break;
    }

    return TRUE;
}


BOOL ProcessListViewMessages(NMHDR* pControlMessage)
{
    LPNMLISTVIEW pListViewMessage;
    wstring strItemPath;
    wstring strObjectName;
    wchar_t wszItemText[MAX_PATH];

    HWND hTreeView;
    HTREEITEM hTreeItem = {};
    TVITEM tvItem;

    HWND hStatic;
    wstring wstrSDDLInformation;

    pListViewMessage = reinterpret_cast<LPNMLISTVIEW>(pControlMessage);

    switch (pControlMessage->code)
    {
    case LVN_ITEMCHANGED:
        if (pListViewMessage->iItem > -1)
        {
            hStatic = GetDlgItem(g_hSecurityDialog, IDC_STATIC_SDDLINFO);
            hTreeView = GetDlgItem(g_hSecurityDialog, IDC_FOLDERS);

            memset(wszItemText, 0, sizeof(wchar_t)*MAX_PATH);
            memset(&tvItem, 0, sizeof(TVITEM));

            //Get the currently select item path from the tree view
            hTreeItem = TreeView_GetSelection(hTreeView);
            tvItem.hItem = hTreeItem;
            tvItem.mask = TVIF_HANDLE;

            TreeView_GetItem(hTreeView, &tvItem);

            TreeViewGetItemPath(hTreeView, &tvItem, &strItemPath);

            //Get the selected item's text from the list view
            ListView_GetItemText(pControlMessage->hwndFrom, pListViewMessage->iItem, pListViewMessage->iSubItem, wszItemText, MAX_PATH);

            strObjectName.assign(strItemPath);
            strObjectName.append(wszItemText);

            CDirectoryManagement::GetObjectSDDLInformation(strObjectName.data(), &wstrSDDLInformation);

            Static_SetText(hStatic, wstrSDDLInformation.data());
        }
        break;
    }

    return TRUE;
}


BOOL ProcessControlMessage(NMHDR* pControlMessageInfo)
{
    switch (pControlMessageInfo->idFrom)
    {
    case IDC_FOLDERS:
        ProcessFolderControlMessage(pControlMessageInfo);
        break;
    case IDC_LIST_FILES:
        ProcessListViewMessages(pControlMessageInfo);
        break;
    }

    return TRUE;
}



//////////////////////////////////////////////////////////////
// SecurityInspectorDialog dialog proc
//
//////////////////////////////////////////////////////////////
INT_PTR CALLBACK SecurityDialogProc(HWND hDlgWindow, UINT uiMessage, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    UNREFERENCED_PARAMETER(wParam);
    
    switch (uiMessage)
    {
    case WM_INITDIALOG:
        g_hSecurityDialog = hDlgWindow;
        LoadDirectoryList(hDlgWindow);
        return static_cast<INT_PTR>(TRUE);
    case WM_NOTIFY:
        ProcessControlMessage(reinterpret_cast<NMHDR*>(lParam));
        break;
    case WM_SIZE:
        ResizeSecurityDialog(hDlgWindow);
        return static_cast<INT_PTR>(TRUE);
    }

    return static_cast<INT_PTR>(FALSE);
}


HTREEITEM TreeViewAddItem(HWND hTreeView, wstring* strItem, HTREEITEM hParentItem)
{
    static int iItem = 0;

    HTREEITEM hTreeItem;
    TVINSERTSTRUCT tvInsert;
    TVITEMEX tvItem;

    memset(&tvInsert, 0, sizeof(TVINSERTSTRUCT));
    memset(&tvItem, 0, sizeof(TVITEM));

    //If hParentItem then modify it to support having children items
    if (hParentItem)
    {
        TVITEMEX tvParentItem;
        memset(&tvParentItem, 0, sizeof(TVITEM));

        //Set the newly inserted item attributes (the text)
        tvParentItem.mask = TVIF_HANDLE;
        tvParentItem.hItem = hParentItem;

        TreeView_GetItem(hTreeView, &tvParentItem);

        //Verify that the item will allow child items
        if (tvParentItem.cChildren == 0)
        {
            tvParentItem.mask |= TVIF_CHILDREN;
            tvParentItem.cChildren = 1;
        }

        TreeView_SetItem(hTreeView, &tvParentItem);
    }

    //Setup and insert a new item
    tvInsert.hParent = hParentItem;
    tvInsert.hInsertAfter = TVI_LAST;
    
    tvInsert.item.mask = TVIF_TEXT;
    tvInsert.item.pszText = const_cast<wchar_t*>(strItem->data());
    tvInsert.item.cchTextMax = static_cast<int>(strItem->length());
  
    hTreeItem = TreeView_InsertItem(hTreeView,&tvInsert);

    return hTreeItem;
}


VOID LoadDirectoryList( HWND hSecurityDialog )
{
    HTREEITEM hDrive;
    HTREEITEM hFolder;
    vector<wstring> vstrDriveList;
    vector<wstring> vstrFolderList;
    vector<wstring> vstrSubFolderList;
    wchar_t strPath[32767];

    //Get the handle to the tree view control
    HWND hTreeView = GetDlgItem(hSecurityDialog, IDC_FOLDERS);

    //Add a root node (computer)
    //wstring strRoot(L"Computer");
    //hRoot = TreeViewAddItem(hTreeView, &strRoot);

    CDirectoryManagement::EnumerateDrives(&vstrDriveList);

    for (unsigned int iDrive = 0; iDrive < vstrDriveList.size(); iDrive++)
    {
        hDrive = TreeViewAddItem(hTreeView, (wstring*)&vstrDriveList[iDrive]);

        //Clear variables
        vstrFolderList.clear();
        memset((void*)strPath, 0, sizeof(wchar_t) * 32767);
        //Create the path to the subdirectory
        _swprintf_p(strPath, 32767, TEXT("%s\\*"), vstrDriveList[iDrive].c_str());

        CDirectoryManagement::EnumerateFolders(strPath, &vstrFolderList);

        for (unsigned int iDir = 0; iDir < vstrFolderList.size(); iDir++)
        {
            if ( (wcscmp(vstrFolderList[iDir].data(), TEXT(".")) == 0) ||
                 (wcscmp(vstrFolderList[iDir].data(), TEXT("..")) == 0) )
            {
                continue;
            }

            hFolder = TreeViewAddItem(hTreeView, (wstring*)&vstrFolderList[iDir],hDrive);

            //Clear variables
            vstrSubFolderList.clear();
            memset((void*)strPath, 0, sizeof(wchar_t) * 32767);
            //Create the path to the subdirectory
            _swprintf_p(strPath, 32767, TEXT("%s\\%s\\*"), vstrDriveList[iDrive].c_str(), vstrFolderList[iDir].c_str());

            //Get and add a subdirectory list
            CDirectoryManagement::EnumerateFolders(strPath, &vstrSubFolderList);
            if (vstrSubFolderList.size() > 2)
            {
                for (unsigned int iSubItem = 0; iSubItem < vstrSubFolderList.size(); iSubItem++)
                {
                    if ( (wcscmp(vstrSubFolderList[iSubItem].data(), TEXT(".")) == 0) ||
                         (wcscmp(vstrSubFolderList[iSubItem].data(), TEXT("..")) == 0) )
                    {
                        continue;
                    }

                    TreeViewAddItem(hTreeView, (wstring*)&vstrSubFolderList[iSubItem], hFolder);
                }
            }
        }
    }
}


HRESULT ListViewClear(HWND hListView)
{
    if (hListView == 0)
    {
        return E_INVALIDARG;
    }

    if (ListView_DeleteAllItems(hListView) == FALSE)
    {
        return E_FAIL;
    }

    return S_OK;
}



HRESULT ListViewAddItems(HWND hListView, vector<wstring>* pstrItems)
{
    if (hListView == 0 || pstrItems == nullptr)
    {
        return E_INVALIDARG;
    }

    for (int iItem = 0; iItem < pstrItems->size(); iItem++)
    {
        if ( (wcscmp((*pstrItems)[iItem].data(), TEXT(".")) == 0) ||
             (wcscmp((*pstrItems)[iItem].data(), TEXT("..")) == 0) )
        {
            continue;
        }


        ListViewAddItem(hListView, (wstring*)&(*pstrItems)[iItem]);
    }

    return S_OK;
}



HRESULT ListViewAddItem(HWND hListView, wstring* strItem)
{
    if (hListView == 0 || strItem == nullptr)
    {
        return E_INVALIDARG;
    }

    LVITEM lvItem = {};

    memset(&lvItem, 0, sizeof(LVITEM));

    lvItem.mask = LVIF_TEXT;
    lvItem.pszText = const_cast<wchar_t*>(strItem->data());
    lvItem.cchTextMax = static_cast<int>(strItem->size());

    if (ListView_InsertItem(hListView, &lvItem) == -1)
    {
        return E_FAIL;
    }

    return S_OK;
}



//////////////////////////////////////////////////////////////////////
// EnumerateDrives
// Output: pDriveList - pointer to a vector list of wstrings

HRESULT CDirectoryManagement::EnumerateDrives(__out vector<wstring>* pDriveList)
{
    wchar_t* pLogicalDriveStrings = nullptr;
    wchar_t* pStringBufferPointer = nullptr;
    wchar_t* pStringBufferBackslash = nullptr;
    wchar_t* pStringBufferEnd = nullptr;

    unsigned long ulLogicalDriveStringsLength = 0;

    unsigned short usStringIndex = 0;

    if (pDriveList == nullptr)
    {
        return E_INVALIDARG;
    }

    //Get the size of the buffer necessary to store all the drive strings
    ulLogicalDriveStringsLength = GetLogicalDriveStringsW(0, nullptr);

    //allocate the buffer
    pLogicalDriveStrings = (wchar_t*)calloc(ulLogicalDriveStringsLength, sizeof(wchar_t));

    if (pLogicalDriveStrings == nullptr)
    {
        return E_OUTOFMEMORY;
    }

    //Get the logical drive name strings
    GetLogicalDriveStringsW(ulLogicalDriveStringsLength, pLogicalDriveStrings);

    //Walk the string buffer and get the list of drives
    pStringBufferEnd = pStringBufferPointer = pLogicalDriveStrings;
    pStringBufferEnd = (wchar_t*)((__int64)pLogicalDriveStrings + (ulLogicalDriveStringsLength * sizeof(wchar_t)) - 2);
    
    while (pStringBufferPointer < pStringBufferEnd)
    {
        //Remove the trailing '\\' character before putting in the vector list
        pStringBufferBackslash = wcschr(pStringBufferPointer, L'\\');

        //Verify that the character was found, otherwise unfortunate things can happen
        if (pStringBufferBackslash != nullptr)
        {
            pStringBufferBackslash[0] = L'\0';
        }

        //Add string to vector list
        pDriveList->push_back(pStringBufferPointer);

        //Increment the pointer past the end of the current string (including null)
        pStringBufferPointer += ((*pDriveList)[usStringIndex].length()) + 2;

        //Increment the current string (and vector list item) index
        usStringIndex++;
    }

    //free the allocated buffer
    free(pLogicalDriveStrings);


    return S_OK;
}


//////////////////////////////////////////////////////////////////////
// EnumerateDirectories
// Input: strPath - system path to retrieve list of directories
// Output: pDirectoryList - pointer to a vector list of wstrings


HRESULT CDirectoryManagement::EnumerateFolders( __in wchar_t* szPath, __out vector<wstring>* pFolderList )
{
    WIN32_FIND_DATA wfdFolderInfo;

    if (pFolderList == nullptr)
    {
        return E_INVALIDARG;
    }

    HANDLE hFindHandle = FindFirstFileExW(szPath,
                                          FindExInfoBasic,
                                          &wfdFolderInfo,
                                          FindExSearchLimitToDirectories,
                                          NULL,
                                          FIND_FIRST_EX_LARGE_FETCH);

    if (hFindHandle == INVALID_HANDLE_VALUE)
    {
        return E_FAIL;
    }

    if (wfdFolderInfo.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
    {
        //OutputDebugStringW(wfdFolderInfo.cFileName);
        //OutputDebugStringW(_T("\r\n"));

        pFolderList->push_back(wfdFolderInfo.cFileName);
    }


    while (FindNextFileW(hFindHandle, &wfdFolderInfo))
    {
        if (wfdFolderInfo.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
        {
            //OutputDebugStringW(wfdFolderInfo.cFileName);
            //OutputDebugStringW(_T("\r\n"));
    
            pFolderList->push_back(wfdFolderInfo.cFileName);
        }
    }

    return S_OK;
}


//////////////////////////////////////////////////////////////////////
// EnumerateDirectories
// Input: strPath - system path to retrieve list of files from
// Output: pFileList - pointer to a vector list of wstrings

HRESULT CDirectoryManagement::EnumerateFiles(__in wchar_t* szPath, __out vector<wstring>* pFileList)
{
    WIN32_FIND_DATA wfdFileInfo;

    if (pFileList == nullptr)
    {
        return E_INVALIDARG;
    }

    HANDLE hFindHandle = FindFirstFileExW(szPath,
                                          FindExInfoBasic,
                                          &wfdFileInfo,
                                          FindExSearchNameMatch,
                                          NULL,
                                          NULL);

    //Validate the find handle
    if (hFindHandle == INVALID_HANDLE_VALUE)
    {
        return E_FAIL;
    }

    //Collect the list of FileNames
    do 
    {
        pFileList->push_back(wfdFileInfo.cFileName);
    } while (FindNextFileW(hFindHandle, &wfdFileInfo));

    return S_OK;
}

HRESULT CDirectoryManagement::GetObjectSDDLInformation(__in const wchar_t* szObjectName, __out wstring* pstrSDDLInformation)
{
    PSECURITY_DESCRIPTOR pObjectSecurityDescriptor = nullptr;
    //Retrieve full set of security information
    SECURITY_INFORMATION siFullPermissions = ATTRIBUTE_SECURITY_INFORMATION | DACL_SECURITY_INFORMATION | SACL_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION;
    SECURITY_INFORMATION siReducedPermissions = ATTRIBUTE_SECURITY_INFORMATION | DACL_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION;
    SECURITY_INFORMATION* pUsedPermissions = &siFullPermissions;
    //siRetrieveInformation |= LABEL_SECURITY_INFORMATION | SCOPE_SECURITY_INFORMATION;
    DWORD dwReturnValue = 0;
    wchar_t* pSDDLString;
    PSID pSidOwner;
    PSID pSidGroup;
    PACL pDacl;
    PACL pSacl;

    if (szObjectName == nullptr || pstrSDDLInformation == nullptr)
    {
        return E_INVALIDARG;
    }

    pstrSDDLInformation->assign(szObjectName);
    pstrSDDLInformation->append(L"\r\n");


    dwReturnValue = GetNamedSecurityInfo(szObjectName, SE_FILE_OBJECT, *pUsedPermissions, &pSidOwner, &pSidGroup, &pDacl, &pSacl, &pObjectSecurityDescriptor);
    
    //If failed, try again with a reduced permission set
    if (dwReturnValue > 0)
    {
        pUsedPermissions = &siReducedPermissions;
        dwReturnValue = GetNamedSecurityInfo(szObjectName, SE_FILE_OBJECT, *pUsedPermissions, &pSidOwner, &pSidGroup, &pDacl, &pSacl, &pObjectSecurityDescriptor);
    }


    //Check for success before converting the Security Descriptor to a string
    if (dwReturnValue == ERROR_SUCCESS)
    {
        ConvertSecurityDescriptorToStringSecurityDescriptorW(pObjectSecurityDescriptor, SDDL_REVISION_1, *pUsedPermissions, &pSDDLString, NULL);

        //Assign to the string for output
        pstrSDDLInformation->append(pSDDLString);

        LocalFree(pObjectSecurityDescriptor);
        LocalFree(pSDDLString);
    }
    else
    {
        pstrSDDLInformation->append(L"Failed to retrieve SDDL information");
        return E_FAIL;
    }

    return S_OK;
}


HRESULT CProcessToken::GetTokenPrivileges(__out FIXED_SIZE_TOKEN_PRIVILEGES* pFixedSizeTokenPrivileges)
{
    HANDLE hProcessToken = INVALID_HANDLE_VALUE;
    FIXED_SIZE_TOKEN_PRIVILEGES fstpProcessTokenPrivileges = {};
    DWORD dwBufferLength = 0;
    DWORD dwLastError = 0;
    BOOL bReturn = 0;

    if (pFixedSizeTokenPrivileges == nullptr)
    {
        return E_INVALIDARG;
    }

    memset(pFixedSizeTokenPrivileges, 0, sizeof(FIXED_SIZE_TOKEN_PRIVILEGES));

    //Retrieve the token and get it's privilege list
    if (FAILED(GetTokenHandle(&hProcessToken)))
    {
        return E_FAIL;
    }

    //Get the list of the token privileges
    bReturn = GetTokenInformation(hProcessToken, TokenPrivileges, pFixedSizeTokenPrivileges, sizeof(FIXED_SIZE_TOKEN_PRIVILEGES), &dwBufferLength);
    
    //Close the Token Handle
    CloseHandle(hProcessToken);

    //Check the return from getting the token privilege list
    if( bReturn == FALSE )
    {
        dwLastError = GetLastError();
        return HRESULT_FROM_WIN32(dwLastError);
    }

    return S_OK;
}


HRESULT CProcessToken::GetTokenHandle( __out PHANDLE pTokenHandle)
{
    HANDLE hProcess = 0;
    DWORD dwLastError = 0;

    hProcess = GetCurrentProcess();

    if (hProcess == 0)
    {
        return E_FAIL;
    }

    if (OpenProcessToken(hProcess, TOKEN_ALL_ACCESS, pTokenHandle) > 0)
    {
        dwLastError = GetLastError();

        return HRESULT_FROM_WIN32(dwLastError);
    }

    return S_OK;
}


HRESULT CProcessToken::DoesTokenHavePrivilege(LPCTSTR tszPrivilegeName)
{
    FIXED_SIZE_TOKEN_PRIVILEGES fstpProcessTokenPrivileges = {};
    DWORD dwLastError = 0;
    LUID luidPrivilege;

    //Get the LUID for the specific privilege constant
    if (LookupPrivilegeValue(NULL, tszPrivilegeName, &luidPrivilege) == FALSE)
    {
        dwLastError = GetLastError();
        return HRESULT_FROM_WIN32(dwLastError);
    }

    //Get the list of privileges for the process token
    if (FAILED(GetTokenPrivileges(&fstpProcessTokenPrivileges)))
    {
        return E_FAIL;
    }

    //Compare the list of the retrieved token privilege list against the specific passed in privilege
    for (DWORD dwIndex = 0; dwIndex < fstpProcessTokenPrivileges.PrivilegeCount; dwIndex++)
    {
        if (luidPrivilege.HighPart == fstpProcessTokenPrivileges.Privileges[dwIndex].Luid.HighPart &&
            luidPrivilege.LowPart == fstpProcessTokenPrivileges.Privileges[dwIndex].Luid.LowPart)
        {
            return S_OK;
        }
    }

    return S_FALSE;
}

HRESULT CProcessToken::EnableTokenPrivilege(LPCTSTR tszPrivilegeName)
{
    HANDLE hProcessToken = INVALID_HANDLE_VALUE;
    FIXED_SIZE_TOKEN_PRIVILEGES fstpProcessTokenPrivileges = {};
    TOKEN_PRIVILEGES tpEnableTokenPrivilege = {};
    DWORD dwLastError = 0;
    BOOL bReturn = 0;

    if (FAILED(DoesTokenHavePrivilege(tszPrivilegeName)))
    {
        return E_FAIL;
    }

    memset(&fstpProcessTokenPrivileges, 0, sizeof(FIXED_SIZE_TOKEN_PRIVILEGES));

    //Get the list of privileges for the process token
    if (FAILED(GetTokenPrivileges(&fstpProcessTokenPrivileges)))
    {
        return E_FAIL;
    }

    for (DWORD dwPrivilege = 0; dwPrivilege < fstpProcessTokenPrivileges.PrivilegeCount; dwPrivilege++)
    {
        fstpProcessTokenPrivileges.Privileges[dwPrivilege].Attributes = SE_PRIVILEGE_ENABLED;
    }

    memset(&tpEnableTokenPrivilege, 0, sizeof(TOKEN_PRIVILEGES));
    tpEnableTokenPrivilege.PrivilegeCount = 1;
    tpEnableTokenPrivilege.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;

    //Get the LUID for the specific privilege constant
    if (LookupPrivilegeValue(NULL, tszPrivilegeName, &(tpEnableTokenPrivilege.Privileges[0].Luid)) == FALSE)
    {
        dwLastError = GetLastError();
        return HRESULT_FROM_WIN32(dwLastError);
    }

    //Retrieve the token and get it's privilege list
    if (FAILED(GetTokenHandle(&hProcessToken)))
    {
        return E_FAIL;
    }

    bReturn = AdjustTokenPrivileges(hProcessToken, FALSE, (PTOKEN_PRIVILEGES)&fstpProcessTokenPrivileges, 0, NULL, NULL);


    //Close the token handle
    CloseHandle(hProcessToken);
        
    if( bReturn == FALSE ) 
    {
        dwLastError = GetLastError();
        return HRESULT_FROM_WIN32(dwLastError);
    }

    return S_OK;
}

