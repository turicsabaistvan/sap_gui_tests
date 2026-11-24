*** Settings ***
Library     Process
Library     ${CURDIR}/SapComLibrary.py


*** Variables ***
${SAPSHCUT_EXE}         C:\\Program Files\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe
${SAP_APP_SERVER}       %{SAP_APP_SERVER=207.175.121.75}
${SAP_SYS_NUM}          %{SAP_SYS_NUM=00}
${SAP_SYS_ID}           %{SAP_SYS_ID=ABA}
${SAP_CLIENT}           %{SAP_CLIENT=800}
${SAP_USER}             %{SAP_USER=IDADMIN}
${SAP_PASS}             %{SAP_PASS=ides123}
${SAP_LANG}             %{SAP_LANG=EN}
${SAP_PROC}             ${NONE}

# Standard IDES SD Master Data
${ORDER_TYPE}           OR
${SALES_ORG}            1000
${DIST_CHANNEL}         10
${DIVISION}             00
${SOLD_TO_PARTY}        1000
${SHIP_TO_PARTY}        1000
${PO_NUMBER}            RF-AUTO-2026
${MATERIAL_CODE}        M-01
${ORDER_QTY}            2


*** Test Cases ***
Create Sales Order And Verify In VA03 And VBAK Table
    [Setup]    Open Clean SAP Session
    [Tags]    test_01
    # 1. Start VA01 and enter organizational data
    Run Transaction    /nVA01
    Input Text    wnd[0]/usr/ctxtVBAK-AUART    ${ORDER_TYPE}
    Input Text    wnd[0]/usr/ctxtVBAK-VKORG    ${SALES_ORG}
    Input Text    wnd[0]/usr/ctxtVBAK-VTWEG    ${DIST_CHANNEL}
    Input Text    wnd[0]/usr/ctxtVBAK-SPART    ${DIVISION}
    Send Vkey    0
    Acknowledge Warnings And Popups

    # 2. Fill Sold-To Party, PO Number, and first line item
    Input Text
    ...    wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR
    ...    ${SOLD_TO_PARTY}
    Input Text
    ...    wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR
    ...    ${SHIP_TO_PARTY}
    Input Text    wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD    ${PO_NUMBER}
    Input Text
    ...    wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]
    ...    ${MATERIAL_CODE}
    Input Text
    ...    wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]
    ...    ${ORDER_QTY}
    Send Vkey    0
    Acknowledge Warnings And Popups

    # 3. Save Sales Order and extract created document number
    Save Document
Test_2
    [Tags]    test_02
    ${status_msg}=    Get Statusbar Message
    Log    Created Order Status: ${status_msg}
    ${order_num}=    Extract Document Number From Statusbar

    # 4. Verify in VA03 (Display Sales Order)
    Run Transaction    /nVA03
    Input Text    wnd[0]/usr/ctxtVBAK-VBELN    ${order_num}
    Send Vkey    0
    Acknowledge Warnings And Popups
    ${displayed_po}=    Get Text    wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD
    Should Be Equal As Strings    ${displayed_po}    ${PO_NUMBER}

    # 5. Verify directly in SAP Database Table VBAK via SE16N
    Run Transaction    /nSE16N
    Input Text    wnd[0]/usr/ctxtGD-TAB    VBAK
    Send Vkey    0
    Input Text    wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]    ${order_num}
    Send Vkey    8
    ${sbar_type}=    Get Statusbar Type
    Should Not Be Equal As Strings    ${sbar_type}    E
    [Teardown]    Close SAP Session


*** Keywords ***
Open Clean SAP Session
    Kill Existing SAP Processes
    ${shcut_proc}=    Start Process    ${SAPSHCUT_EXE}
    ...    -guiparm\=/H/${SAP_APP_SERVER}/S/32${SAP_SYS_NUM}
    ...    -system\=${SAP_SYS_ID}
    ...    -client\=${SAP_CLIENT}
    ...    -user\=${SAP_USER}
    ...    -pw\=${SAP_PASS}
    ...    -language\=${SAP_LANG}
    Set Suite Variable    ${SAP_PROC}    ${shcut_proc}
    Connect To Active Session    timeout=30
    Acknowledge Warnings And Popups
    Element Should Be Present    wnd[0]/tbar[0]/okcd

Close SAP Session
    Run Keyword And Ignore Error    Run Transaction    /nex
    Run Keyword And Ignore Error    Terminate Process    ${SAP_PROC}    kill=True
    Run Keyword And Ignore Error    Kill Existing SAP Processes
