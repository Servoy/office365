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

				// parse the servoy attachments
				var paramAttachments = [];
				for (var i = 0; attachments && i < attachments.length; i++) {
					var attachment = attachments[i]
					var paramAttachment = {
						type: attachment.type,
						name: attachment.name,
						itemId: attachment.itemId,
						isInline: attachment.isInline ? true : false,
						url: attachment.url//attachment.media ? 'http://localhost:8080/' + attachment.media : null
					}

					paramAttachments.push(paramAttachment);
				}

				try {
					var params = {
						toRecipients: parent.deepCopyArray(toRecipients), // Copy the To line from current item
						ccRecipients: parent.deepCopyArray(ccRecipients), // Copy the To line from current item
						subject: subject,
						htmlBody: htmlBody
					}

					// if there are attachments
					if (paramAttachments && paramAttachments.length) {
						params.attachments = parent.deepCopyArray(paramAttachments);
					}

					Office.context.mailbox.displayNewMessageForm(params);

					officeResultDeferred.resolve(JSON.stringify(Office.context.mailbox.item.to));

				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},

			addRecipients: function(recipients, onError) {

				var officeResultDeferred = $q.defer();
				try {

					Office.context.mailbox.item.to.addAsync(['paronne@servoy.com'], {test: "sample"}, function(result) {
							if (result.error) {
								if (onError) {
									resolveError(JSON.stringify(result.error), onError, officeResultDeferred);
								}
								officeResultDeferred.resolve(false);
							} else {
								officeResultDeferred.resolve(true);
							}
						});

				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			
			addRecipients: function(recipients, onError) {
				var officeResultDeferred = $q.defer();

				recipients = parent.deepCopyArray(recipients);

				try {
					Office.context.mailbox.item.to.addAsync(recipients, function(result) {

							if (result.error) {
								if (onError) {
									resolveError(JSON.stringify(result.error), onError, officeResultDeferred);
								}
								officeResultDeferred.resolve(false);
							} else {
								officeResultDeferred.resolve(true);
							}
						});

				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			
			addCCRecipients: function(recipients, onError) {

				var officeResultDeferred = $q.defer();

				recipients = parent.deepCopyArray(recipients);

				try {

					Office.context.mailbox.item.cc.addAsync(recipients, function(result) {

							if (result.error) {
								if (onError) {
									resolveError(JSON.stringify(result.error), onError, officeResultDeferred);
								}
								officeResultDeferred.resolve(false);
							} else {
								officeResultDeferred.resolve(true);
							}
						});

				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},

			addBCCRecipients: function(recipients, onError) {

				var officeResultDeferred = $q.defer();

				recipients = parent.deepCopyArray(recipients);

				try {

					Office.context.mailbox.item.bcc.addAsync(recipients, function(result) {
							if (result.error) {
								if (onError) {
									resolveError(JSON.stringify(result.error), onError, officeResultDeferred);
								}
								officeResultDeferred.resolve(false);
							} else {
								officeResultDeferred.resolve(true);
							}
						});

				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},

			getSelectedMessageBodyType: function(onError) {

				var officeResultDeferred = $q.defer();

				try {
					var message = Office.context.mailbox.item;

					message.body.getTypeAsync({ }, function(result) {
							if (result.error) {
								resolveError(result.error, onError, officeResultDeferred);
								officeResultDeferred.resolve(false);
							} else {
								officeResultDeferred.resolve(result);
							}
						})

				} catch (e) {
					resolveError(e, onError, officeResultDeferred);
				}

				return officeResultDeferred.promise;
			},

			getSelectedMessageBody: function(coercionType, onError) {

				var officeResultDeferred = $q.defer();

				try {
					coercionType = typeof coercionType === undefined || coercionType == 'html' ? /*Office.CoercionType.Text*/ 'text' : coercionType;

					var message = Office.context.mailbox.item;
					if (typeof message.body.getAsync === "function") {
						var bodyText = message.body.getAsync(coercionType, { }, function(result) {
								if (result.error) {
									resolveError(result.error, onError, officeResultDeferred)
									officeResultDeferred.resolve(false)
								} else {
									officeResultDeferred.resolve(result)
								}
							})
					} else {
						officeResultDeferred.resolve(message.body)
					}
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},

			getSelectedMessageSubject: function(onError) {
				var officeResultDeferred = $q.defer();

				try {

					var message = Office.context.mailbox.item;

					if (typeof message.subject.getAsync === "function") {
						message.subject.getAsync(function(result) {
							if (result.error) {
								officeResultDeferred.resolve(false);
							} else {
								officeResultDeferred.resolve(result);
							}
						});
					} else {
						officeResultDeferred.resolve(message.subject)
					}

				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},

			setSubject: function(subject, onError) {
				var officeResultDeferred = $q.defer();

				try {
					Office.context.mailbox.item.subject.setAsync(subject, function(result) {
							if (result.error) {
								resolveError(result.error, onError, officeResultDeferred)
								officeResultDeferred.resolve(false);
							} else {
								officeResultDeferred.resolve(true);
							}
						})
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},

			setBodyText: function(text, coercionType, onError) {

				var officeResultDeferred = $q.defer();

				try {

					coercionType = typeof coercionType === 'undefined' ? Office.CoercionType.Text : coercionType;

					var messageBody = Office.context.mailbox.item.body;

					messageBody.setAsync(text, { coercionType: coercionType }, function(result) {

							if (result.error) {
								resolveError(result.error, onError, officeResultDeferred)
								officeResultDeferred.resolve(false);
							} else {
								officeResultDeferred.resolve(true);
							}
						});

				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			
			prependBodyText: function(text, coercionType, onError) {

				var officeResultDeferred = $q.defer();

				try {

					coercionType = typeof coercionType === 'undefined' ? Office.CoercionType.Text : coercionType;

					var messageBody = Office.context.mailbox.item.body;

					messageBody.prependAsync(text, { coercionType: coercionType }, function(result) {

							if (result.error) {
								resolveError(result.error, onError, officeResultDeferred);
								officeResultDeferred.resolve(false);
							} else {
								officeResultDeferred.resolve(true);
							}

						});

				} catch (e) {
					resolveError(e, onError, officeResultDeferred);
				}

				return officeResultDeferred.promise;
			},

			setSelectedText: function(text, coercionType, onError) {

				var officeResultDeferred = $q.defer();

				try {

					coercionType = typeof coercionType === 'undefined' ? Office.CoercionType.Text : coercionType;

					var message = Office.context.mailbox.item;

					message.setSelectedDataAsync(text, { coercionType: coercionType }, function(result) {

							if (result.error) {
								resolveError(result.error, onError, officeResultDeferred);
								officeResultDeferred.resolve(false);
							} else {
								officeResultDeferred.resolve(true);
							}

						})
				} catch (e) {
					resolveError(e, onError, officeResultDeferred);
				}

				return officeResultDeferred.promise;
			},

			addAttachment: function(file, attachmentName, onError) {

				var officeResultDeferred = $q.defer();

				var base64regex = /^([0-9a-zA-Z+\/]{4})*(([0-9a-zA-Z+\/]{2}==)|([0-9a-zA-Z+\/]{3}=))?$/;

				try {

					var message = Office.context.mailbox.item;

					var _bIsBase64 = base64regex.test(file);
					if (_bIsBase64) {
						message.addFileAttachmentFromBase64Async(file, attachmentName, {}, function(result) {
								if (result.error) {
									resolveError(result.error, onError, officeResultDeferred);
									officeResultDeferred.resolve(false);
								} else {
									officeResultDeferred.resolve(true);
								}
							});
					} else {
						message.addFileAttachmentAsync(file, attachmentName, {}, function(result) {
								if (result.error) {
									resolveError(result.error, onError, officeResultDeferred);
									officeResultDeferred.resolve(false);
								} else {
									officeResultDeferred.resolve(true);
								}
							});
					}
				} catch (e) {
					resolveError(e, onError, officeResultDeferred);
				}

				return officeResultDeferred.promise;
			},
			getAttachments: function(onError) {

				var officeResultDeferred = $q.defer();

				try {
					var result = []

					var message = Office.context.mailbox.item;

					if (typeof message.getAttachmentsAsync === "function") {
						message.getAttachmentsAsync(function(result) {
							if (result.error) {
								resolveError(result.error, onError, officeResultDeferred)
							} else {
								var obj = new Object();
								for (var i = 0; i < result.value.length; i++) {
									var obj = new Object();
									for (var key in result.value[i]) {
										obj[key] = result.value[i][key];
									}
									result.push(obj);
								}
								officeResultDeferred.resolve(result);
							}
						})
					} else {
						for (var i = 0; i < message.attachments.length; i++) {
							var attachment = {};
							for (var key in message.attachments[i]) {
								attachment[key] = message.attachments[i][key];
							}
							result.push(attachment);
						}
						officeResultDeferred.resolve(result);
					}
				} catch (e) {
					resolveError(e, onError, officeResultDeferred);
				}

				return officeResultDeferred.promise;
			},

			getAttachmentContent: function(attachmentId, onError) {

				var officeResultDeferred = $q.defer();

				try {

					var message = Office.context.mailbox.item;

					message.getAttachmentContentAsync(attachmentId, { }, function(result) {
							if (result.error) {
								resolveError(result.error, onError, officeResultDeferred);
							} else {
								officeResultDeferred.resolve(result);
							}
						});

				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			}
		}
	}]).run(function($rootScope, $services, $log) { })