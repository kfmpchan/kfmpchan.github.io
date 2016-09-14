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

var varTimeoutDelay = 1000; // Every command takes an artifical 5 seconds

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
	globalEvent = event;
	setGlobals("icon3", "MailReadUILess1", "Mail Read UiLess 1 Has Finished");
	setTimeout(asyncCallbackFinish, varTimeoutDelay);
}

function uiLessFunctionMailRead2(event)
{
	globalEvent = event;




	Office.context.mailbox.item.loadCustomPropertiesAsync(function (asyncResult) {
  
		if (asyncResult.status == "failed") 
		{
    
			showMessage("Action failed with error: " + asyncResult.error.message);
  
		}
  
		else 
		{
    
		var cusProps = asyncResult.value;
    
    
		cusProps.set("var1", "Doctor Doom");
    
		cusProps.saveAsync();
    
		var out = cusProps.get("var1");
  
		setGlobals("icon4", "MailReadUILess2", out);

		setTimeout(asyncCallbackFinish, varTimeoutDelay);

		}

		
});
}


// ------ MAIL COMPOSE --------

function uiLessFunctionMailCompose1(event)
{
	globalEvent = event;
	setGlobals("icon7", "MailComposeUILess133", "Mail Compose UiLess 1 Has Finished");

	setTimeout(asyncCallbackFinish, varTimeoutDelay);
}

function uiLessFunctionMailCompose2TEST(event)
{
	globalEvent = event;
	var outString = "";

	Office.context.mailbox.item.getSelectedDataAsync("text", 
		function (asyncResult) 
		{
  
			if (asyncResult.status == "failed") 
			{
    
			debugger;
			}
  
			else {
    
				outString = asyncResult.value.data;

				setGlobals("icon8", "MailComposeUILess2", outString);
				setTimeout(asyncCallbackFinish, varTimeoutDelay);
			}
		
});
	

}


function makeEmail() 
{ 
	var strValues="abcdefg12345"; 
	var strEmail = ""; 
	var strTmp; 

	for (var i=0;i<10;i++) 
	{ 
		strTmp = strValues.charAt(Math.round(strValues.length*Math.random())); 
		strEmail = strEmail + strTmp; 
	} 

	strTmp = ""; 
	strEmail = strEmail + "@batbeyond.com"; 
	return strEmail; 
}


function Add5Recipients(recipCollection) 
{
	var JsonArray = {}; 

	
	JsonArray = [];

	for(var i = 0; i < 5; i++)
	{
		JsonArray.push({});
		JsonArray[i]['displayName'] = "Zombie Tim" + i;
		JsonArray[i]['emailAddress'] = makeEmail();
	}	
	recipCollection.setAsync(JsonArray,         
		function (asyncResult) {
            			if (asyncResult.status == Office.AsyncResultStatus.Failed)
							{
                			handleErrors(asyncResult.error.message);
            				}
		});
}

function recipComposeHelper()
{
	var toVar;
	var ccVar;
	var _Item = Office.context.mailbox.item;

	if (_Item.to)
	{
		toVar = _Item.to;
		ccVar = _Item.cc;
	}
	else
	{
		toVar = _Item.requiredAttendees;
		ccVar = _Item.optionalAttendees;
	}

	Add5Recipients(toVar);
	Add5Recipients(ccVar);

}

var itemId="AAMkADhiZDNlNjI4LTE4YzYtNDI0NS1hZTg1LWY0ZGRhYjM2YTg3OQBGAAAAAAB5OY8dgrsJQIeA1SRySophBwCGuw0cS37SQrpvUCTC4feLAAAAAAEMAACGuw0cS37SQrpvUCTC4feLAAAPuAOfAAA=";
  
  
  
function callback() {}





function addAttachment() {
  
	// The values in asyncContext can be accessed in the callback
  
	var options = { 'asyncContext': { var1: 1, var2: 2 } };
  
  
	var attachmentURL = "http://i.imgur.com/ntq8j1J.jpg";
 
	 Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
    	//Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);


}

function uiLessFunctionMailCompose2(event)
{
	setGlobals("icon8", "MailComposeUILess2", "Hi MOm");

	globalEvent = event;


	addAttachment();

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
	setGlobals("icon8", "ApptComposeUILess2", "Appt Compose UiLess 2 Has Finished");

Office.context.mailbox.item.subject.setAsync("New subject! From UILess");

	globalEvent = event;
	setTimeout(asyncCallbackFinish, varTimeoutDelay);

}








