var Port = null;
var HostName = "astar.ics.nativemessage.extension"
function connect() {
    if (null == Port) 
    {
        Port = chrome.runtime.connectNative(HostName);
        if (null != Port)
            Port.onDisconnect.addListener(onDisconnected);
    }

}
function vvTab(tab) {
    //alert("OnActiveChanged:" + tab.url);
    connect();
    if (null != Port)
        Port.postMessage({ "text": tab.url });
}
function vvTab2(tab) {
    //alert("OnActiveChanged:" + tab.url);
    connect();
    if (null != Port)
        Port.postMessage({ "text": tab.url });
}
function OnURLUpdate(tabId, changeInfo, tab) {
    //alert("URL Change:" + tab.url);
    connect();
    if (null != Port)
        Port.postMessage({ "text": tab.url });
}
function OnActiveChanged(activeInfo) {

    chrome.tabs.get(activeInfo.tabId, vvTab);

}
function onWindowsFocusChanged(windowId)
{
    chrome.tabs.getSelected(windowId,vvTab2);
}
function onDisconnected() {
    Port = null;
}


chrome.tabs.onUpdated.addListener(OnURLUpdate);
chrome.tabs.onActivated.addListener(OnActiveChanged);
chrome.windows.onFocusChanged.addListener(onWindowsFocusChanged);