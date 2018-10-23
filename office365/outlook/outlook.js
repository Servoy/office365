angular.module('office365Outlook', ['servoy']).factory("office365Outlook", ['$services', '$window', '$log', '$q', function($services, $window, $log, $q) {

		var scope = $services.getServiceScope('office365Outlook');
		var Office = parent.Office;
		var Outlook = parent.Outlook;

		var CONTENT_TYPE = {
			TEXT: "text",
			HTML: "html",
			OOXML: "ooxml"
		}

		var INSERT_LOCATION = { }

		// does it make sense ?
		if (Outlook) {
			INSERT_LOCATION = {
				END: Outlook.InsertLocation.end,
				START: Outlook.InsertLocation.start,
				REPLACE: Outlook.InsertLocation.replace
			}
		}

		/**
		 * @param error
		 * @param onError
		 * @param officeResultDeferred
		 *
		 * resolve promis with error
		 *  */
		function resolveError(error, onError, officeResultDeferred) {
			var msg = typeof (error) === "string" ? error : error.message;
			$log.error('office365Outlook: ', msg);
			if (onError) {
				$window.executeInlineScript(onError.formname, onError.script, [msg]);
			}
			officeResultDeferred.resolve(null);
		}

		return {
			displayNewMessageForm: function(toRecipients, ccRecipients, subject, htmlBody, attachments, onError) {

				var officeResultDeferred = $q.defer();
//
//				if (!Office.context.requirements.isSetSupported('Mailbox', '1.6')) {
//					//resolveError("Sorry, this add-in only works with newer versions of Outlook.", onError, officeResultDeferred);
//					$log.error("Sorry, this add-in only works with newer versions of Outlook.");
//				} else {
//					//resolveError("ok", onError, officeResultDeferred);
//					$log.error("yes initialized")
//				}

				// toRecipients: [{'emailAddress':'test@com.com','displayName':'Paolo Aronne','recipientType':'externalUser'}];

				// parse the servoy attachments
				var paramAttachments = [];
				for (var i = 0; attachments && i < attachments.length; i++) {
					var attachment = attachments[i]
					var paramAttachment = {
						type: attachment.type,
						name: attachment.name,
						itemId: attachment.itemId,
						isInline: attachment.isInline ? true : false,
						url: attachment.media ? 'http://localhost:8080/' + attachment.media : null
					}

					paramAttachments.push(paramAttachment);
				}

				try {
					var params = {
						toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
						ccRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
						subject: subject,
						htmlBody: htmlBody
					}
					
					// if there are attachments
					if (paramAttachments && paramAttachments.length) {
						params.attachments = paramAttachments;
						
						//resolveError(paramAttachments[0].url, onError, officeResultDeferred);
						//return;
					}
					
					Office.context.mailbox.displayNewMessageForm(params);

					officeResultDeferred.resolve(JSON.stringify(Office.context.mailbox.item.to));

					function getSelectedDataCallback(result) {
						if (result.status === Office.AsyncResultStatus.Succeeded) {
							officeResultDeferred.resolve(result.value);
						} else {
							resolveError(result.error, onError, officeResultDeferred);
						}
					}
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			getSelectedMessageType: function(options, callback, onError) {

				var officeResultDeferred = $q.defer();

				Outlook.run(function(ctx) {

					var message = ctx.mailbox.item;

					var messageType = message.body.getTypeAsync(options, callback)

					return ctx.sync().then(function() {
						officeResultDeferred.resolve(messageType);
					});
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				});

				return officeResultDeferred.promise;
			},

			getSelectedMessageBody: function(coercionType, options, callback, onError) {

				var officeResultDeferred = $q.defer();
				Outlook.run(function(ctx) {

					coercionType = typeof coercionType === 'undefined' ? Office.CoercionType.Text : coercionType;

					var message = ctx.mailbox.item;

					var bodyText = message.body.getAsync(coercionType, options, callback)

					return ctx.sync().then(function() {
						officeResultDeferred.resolve(message);
					});
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				});

				return officeResultDeferred.promise;
			},

			getSelectedMessageSubject: function() {
				var officeResultDeferred = $q.defer();
				Outlook.run(function(ctx) {

					coercionType = typeof coercionType === 'undefined' ? Office.CoercionType.Text : coercionType;

					var message = ctx.mailbox.item;

					var messageSubject = ctx.mailbox.item

					return ctx.sync().then(function() {
						officeResultDeferred.resolve(messageSubject);
					});
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				});

				return officeResultDeferred.promise;
			},

			setSubject: function() { },

			setBodyText: function(data, options, callback, onError) {

				var officeResultDeferred = $q.defer();
				Outlook.run(function(ctx) {

					var messageBody = ctx.mailbox.item.body;

					messageBody.setAsync(data, options, callback);

					return ctx.sync().then(function() {
						officeResultDeferred.resolve(true);
					})
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				});
			}
		}
	}]).run(function($rootScope, $services, $log) {

	//		Office.onReady()
	//	    .then(function() {
	//	        if (!Office.context.requirements.isSetSupported('OutlookApi', '1.6')) {
	//	            $log.warn("Sorry, this add-in only works with newer versions of Outlook.");
	//	        } else {
	//	        	$log.warn("yes initialized")
	//	        }
	//	    });

})