/**
 * @type {String}
 *
 * @properties={typeid:35,uuid:"1F99DD64-82EC-4BA0-9DCD-9D103C5F2298"}
 */
var subject = 'Hello Outlook';

/**
 * @type {String}
 *
 * @properties={typeid:35,uuid:"1159E8E8-A182-4111-AC68-262E40203D86"}
 */
var toReceipients = null;

/**
 * @type {String}
 *
 * @properties={typeid:35,uuid:"28A433E0-C7BE-4F4A-AA93-55985CF03296"}
 */
var htmlBody = 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>'

/**
 * @properties={typeid:35,uuid:"31B654CA-22B3-4665-BBB0-AAC12C35DAD9",variableType:-4}
 */
var $supplierid = null;

/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"248727EB-526F-4348-ABE4-80012C0FB340"}
 */
function sendMessage(event) {
	//plugins.svyToastr.info("Your message has been sent with success to " + foundset.contactname)
}

/**
 * @param {JSEvent} event the event that triggered the action
 *
 * @properties={typeid:24,uuid:"03800758-904B-4DC1-9A1E-AFCBF21B322D"}
 */
function insertText(event) {
	if (foundset.getSize() > 0) {
		plugins.office365Word.setSelectedData(foundset.getAddressAsText(), scopes.office365.COERCION_TYPE.TEXT, onError);
	}
}

/**
 * @param e
 *
 * @properties={typeid:24,uuid:"CE6793AC-83AA-44AD-84FF-C673576090A5"}
 */
function onError(e) {
	application.output(e)
}

/**

 * @protected
 *
 * @properties={typeid:24,uuid:"7C9E8E2C-627E-460A-B4BE-4EABD599BE19"}
 */
function onSwipeLeft() {
	foundset.setSelectedIndex(foundset.getSelectedIndex() + 1)
}

/**
 * @protected 
 * 
 * @properties={typeid:24,uuid:"E3C07CE3-25E1-4FAF-A329-0D15EBD5A5AF"}
 */
function onSwipeRight() {
	foundset.setSelectedIndex(foundset.getSelectedIndex() - 1)
}

/**
 * Handle changed data, return false if the value should not be accepted. In NGClient you can return also a (i18n) string, instead of false, which will be shown as a tooltip.
 *
 * @param oldValue old value
 * @param newValue new value
 * @param {JSEvent} event the event that triggered the action
 *
 * @return {Boolean}
 *
 * @properties={typeid:24,uuid:"0C4292FC-146F-41CE-898A-9E128DF503BE"}
 * @AllowToRunInFind
 */
function onDataChangeSupplier(oldValue, newValue, event) {
	if ($supplierid) {
		foundset.find();
		foundset.supplierid = $supplierid;
		foundset.search();
	}
	// $supplierid = null;
	return true;
}

/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @properties={typeid:24,uuid:"5AB0381D-8578-4716-B66A-411C1934A852"}
 */
function onComposeEmail(event) {
	/** @type {office365-outlook.attachment} */
	var attachment= new Object();
	attachment.isInline = false;
	attachment.name = "test.png";
	attachment.type = plugins.office365Outlook.ATTACHMENT_TYPE.FILE;
	//attachment.itemId = "testimage";
	attachment.media = "media:///test.png";
	var result = plugins.office365Outlook.displayNewMessageForm(toReceipients.split("\n"),null,subject,htmlBody, [attachment], onError)
	if (result) {
		application.output(result)
	}
}

/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"3BAB3CA7-76C1-4751-ADA7-09E7B3E1348D"}
 */
function btnCancelSearchField(event) {
	$supplierid = null;
}

/**

 * @protected
 *
 * @properties={typeid:24,uuid:"869E08BF-B4AD-4E27-95A7-51D30FF747CD"}
 */
function onAddRecepients() {
	
	/** @type {Array<office365-outlook.recipient>} */
//	var recepients = [{displayName: "test", emailAddress: "paronne@servoy.com"}]
	
	var result = plugins.office365Outlook.setSubject('test',onError)
	if (result) {
		application.output(result)
	}
}
