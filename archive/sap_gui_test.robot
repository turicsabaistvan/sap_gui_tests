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


*** Test Cases ***
SAP GUI Direct Login Logout And Close
    ${shcut_proc}=    Start Process    ${SAPSHCUT_EXE}
    ...    -guiparm\=/H/${SAP_APP_SERVER}/S/32${SAP_SYS_NUM}
    ...    -system\=${SAP_SYS_ID}
    ...    -client\=${SAP_CLIENT}
    ...    -user\=${SAP_USER}
    ...    -pw\=${SAP_PASS}
    ...    -language\=${SAP_LANG}

    Connect To Active Session    timeout=30
    Dismiss Popup If Present
    Element Should Be Present    wnd[0]/tbar[0]/okcd

    # Example: open SE16N and return to Easy Access
    Run Transaction    /nSE16N
    Element Should Be Present    wnd[0]/usr/ctxtGD-TAB

    # Clean logout and terminate
    Run Transaction    /nex
    Terminate Process    ${shcut_proc}    kill=True
