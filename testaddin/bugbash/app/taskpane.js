// Microsoft.OWE.Outlook.V15.initialize("df414b9e-4ba3-4955-9a35-200ab41fcb51", onInitializeComplete);

var replyAction = null;
var itemId = null;
var state = null;
var init = false;
var bodyload = false;
var _UserProfile = null;
var _Om;
var _Item;

Office.initialize = function (reason) 
{

    _Om = Office.context.mailbox;
	_Item = _Om.item;

	//_UserProfile = _Om.userProfile;
    //_settings = context.get_settings();
    //_systemSettings = context.get_systemSettings();

    onInitializeComplete();
}

function showMessage(htmlToShow)
{
	var messagesDiv = document.getElementById('messages');
	messagesDiv.innerHTML = htmlToShow;
}

function onInitializeComplete() 
{
	init = true;
	if (bodyload && init) 
		populate();
}

function loadcomplete() 
{
	bodyload = true;
	if (bodyload && init) 
		populate();
}

function populate() 
{
	showMessage("Ready!");
}

function doAction() 
{
	var mySubject = Office.context.mailbox.item.subject;

	Office.context.mailbox.item.displayReplyForm("<b>Thanks</b> for your message about: " + mySubject + ".<BR> this is my reply...");
}

function doSetSubject()
{
	

Office.context.mailbox.item.subject.setAsync("New subject!"); 
}
