//
var welcomeTitle = "It's time to protect your account!";
var welcomeMessage = "Welcome to the registration.";
var thankyouTitle = "Your account is registered!";
var thankyouMessage = "Thank you for registering.";

//
var msalConfig = {
    auth: {
        clientId: "Enter_the_Application_Id_here", 
        authority: "https://login.microsoftonline.com/Enter_the_Tenant_Info_Here"
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true
    }
};

// this can be used for login or token request, however in more complex situations
// this can have diverging options
var requestObj = {
    scopes: ["openid profile"]
};

var myMSALObj = new Msal.UserAgentApplication(msalConfig);
// Register Callbacks for redirect flow
myMSALObj.handleRedirectCallback(authRedirectCallBack);

function closeBrowser() {
    window.open('','_parent','');
    window.close();
}

function signIn() {
    myMSALObj.loginPopup(requestObj).then(function (loginResponse) {
        //Login Success
        showThankyouMessage();
        //acquireTokenPopupAndCallMSGraph();
    }).catch(function (error) {
        console.log(error);
    });
}

function signOut() {
    myMSALObj.logout();
}

function showWelcomeMessage() {
    var divPageTitle = document.getElementById('pageTitle');
    divPageTitle.innerHTML = welcomeTitle;
    var divMessage = document.getElementById('pageMessage');
    divMessage.innerHTML = welcomeMessage;
}

function showThankyouMessage() {
    var divPageTitle = document.getElementById('pageTitle');
    divPageTitle.innerHTML = thankyouTitle;
    var divUserInfo = document.getElementById('userInfo');
    divUserInfo.innerHTML = myMSALObj.getAccount().name + " (" + myMSALObj.getAccount().userName + ")";
    var divMessage = document.getElementById('pageMessage');
    divMessage.innerHTML = thankyouMessage;
    var buttonClose = document.getElementById('close');
    buttonClose.style = "display: none;";
    var buttonLogin = document.getElementById('signIn');
    buttonLogin.innerHTML = "Sign Out";
    buttonLogin.setAttribute('onclick', 'signOut();');
}

function authRedirectCallBack(error, response) {
    if (error) {
        console.log(error);
    }
    else {
        if (response.tokenType === "access_token") {
            callMSGraph(graphConfig.graphEndpoint, response.accessToken, graphAPICallback);
        } else {
            console.log("token type is:" + response.tokenType);
        }
    }
}

function requiresInteraction(errorCode) {
    if (!errorCode || !errorCode.length) {
        return false;
    }
    return errorCode === "consent_required" ||
        errorCode === "interaction_required" ||
        errorCode === "login_required";
}


// Browser check variables
var ua = window.navigator.userAgent;
var msie = ua.indexOf('MSIE ');
var msie11 = ua.indexOf('Trident/');
var msedge = ua.indexOf('Edge/');
var isIE = msie > 0 || msie11 > 0;
var isEdge = msedge > 0;
//If you support IE, our recommendation is that you sign-in using Redirect APIs
//If you as a developer are testing using Edge InPrivate mode, please add "isEdge" to the if check
// can change this to default an experience outside browser use
//var loginType = isIE ? "REDIRECT" : "POPUP";
//
//Forcing redirect login for a better user experience
var loginType = "REDIRECT"

if (loginType === 'POPUP') {
    if (myMSALObj.getAccount()) {// avoid duplicate code execution on page load in case of iframe and popup window.
        showThankyouMessage();
    }
    else {
        showWelcomeMessage();
    }
}
else if (loginType === 'REDIRECT') {
    document.getElementById("signIn").onclick = function () {
        myMSALObj.loginRedirect(requestObj);
    };
    if (myMSALObj.getAccount() && !myMSALObj.isCallback(window.location.hash)) {// avoid duplicate code execution on page load in case of iframe and popup window.
        showThankyouMessage();
    }
    else {
        showWelcomeMessage();
    }
} else {
    console.error('Please set a valid login type');
}
