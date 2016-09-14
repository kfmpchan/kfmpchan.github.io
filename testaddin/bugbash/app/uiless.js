// Microsoft.OWE.Outlook.V15.initialize("df414b9e-4ba3-4955-9a35-200ab41fcb51", onInitializeComplete);



var composeEvent;
var readEvent;
var globalEvent;

Office.initialize = function (reason) 
{
   //debugger;

    _Om = Office.context.mailbox;
	_Item = _Om.item;

	//_UserProfile = _Om.userProfile;
    //_settings = context.get_settings();
    //_systemSettings = context.get_systemSettings();

    //onInitializeComplete();
}

var varTimeoutDelay = 20000; // Every command takes an artifical 5 seconds

function setGlobals(icon, key, message)
{
	globalIcon = icon;
	globalKey = key;
	globalString = message;
}

function asyncCallbackFinish(asyncResult) 
{
	Office.context.mailbox.item.notificationMessages.addAsync(
globalKey, 
        { type: "informationalMessage", icon:  globalIcon, 
          message: globalString, 
          persistent: true
        });
	globalEvent.completed(true);
}

// ------ MAIL READ --------

function uiLessFunctionMailRead1(event)
{
	// debugger;
	var settingsValue = Office.context.roamingSettings.get("mySetting");
	
	globalEvent = event;
	setGlobals("icon3", "MailReadUILess1", "Mail Read UiLess 1 Has Finished. mySetting value is: " + settingsValue);
	setTimeout(asyncCallbackFinish, varTimeoutDelay);
}

function uiLessFunctionMailRead2(event)
{

 //I have bugged this, to test timeouts

//	globalEvent = event;
//	setGlobals("icon4", "MailReadUILess2", "Mail Read UiLess 2 Has Finished");
//	setTimeout(asyncCallbackFinish, varTimeoutDelay);
}


// ------ MAIL COMPOSE --------

function uiLessFunctionMailCompose1(event)
{
	globalEvent = event;
	setGlobals("icon7", "MailComposeUILess133", "Mail Compose UiLess 1 Has Finished");

	setTimeout(asyncCallbackFinish, varTimeoutDelay);
}

function uiLessFunctionMailCompose2(event)
{
	globalEvent = event;


	setGlobals("icon8", "MailComposeUILess2", "Mail Compose UiLess 2 Has Finished");
	setTimeout(asyncCallbackFinish, varTimeoutDelay);
}

// ------ APPT ReadAttendee --------

function uiLessFunctionApptAttendee1(event)
{
	setGlobals("icon3", "ApptReadUILess1", "Appt Read UiLess 1 Has Finished");
	globalEvent = event;
	setTimeout(asyncCallbackFinish, varTimeoutDelay);

}

function uiLessFunctionApptAttendee2(event)
{
	setGlobals("icon4", "ApptReadUILess2", "Appt Read UiLess 2 Has Finished");
	globalEvent = event;
	setTimeout(asyncCallbackFinish, varTimeoutDelay);

}


// ------ APPT ComposeOrganizer--------


function uiLessFunctionApptOrganizer1(event)
{
	setGlobals("icon7", "ApptComposeUILess1", "Appt Compose UiLess 1 Has Finished");

	globalEvent = event;
	setTimeout(asyncCallbackFinish, varTimeoutDelay);
}


function uiLessFunctionApptOrganizer2(event)
{
	setGlobals("icon8", "ApptComposeUILess2", "Appt Compose UiLess 1 Has Finished");

	globalEvent = event;
	setTimeout(asyncCallbackFinish, varTimeoutDelay);

}








