angular.module('office365Word', ['servoy']).factory("office365Word", ['$services', '$window', '$log', '$q', function($services, $window, $log, $q) {

		// MANIFEST LINE: JS-ClientLibs: //appsforoffice.microsoft.com/lib/1/hosted/office.debug.js;version=2.2;name=office.js
		// JS-ClientLibs: office365/lib/office.js;version=2.2;name=office.js,office365/lib/MicrosoftAjax.js;version=2.2;name=MicrosoftAjax.js,office365/lib/o15apptofilemappingtable.js;version=2.2;name=o15apptofilemappingtable.js

		var scope = $services.getServiceScope('office365Word');
		var Office = parent.Office;
		var Word = parent.Word;

		var CONTENT_TYPE = {
			TEXT: "text",
			HTML: "html",
			OOXML: "ooxml"
		}

		var PARTS = {
			BODY: "body",
			PARAGRAPH: "paragraphs"
		}

		var INSERT_LOCATION = {}
		
		// does it make sense ?
		if (Word) {
			INSERT_LOCATION = {
				END: Word.InsertLocation.end,
				START: Word.InsertLocation.start,
				REPLACE: Word.InsertLocation.replace
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
			$log.error('office365Word: ', msg);
			if (onError) {
				$window.executeInlineScript(onError.formname, onError.script, [msg]);
			}
			officeResultDeferred.resolve(null);
		}

		/**
		 * @param {String} title
		 * @param {Function} syncCallback
		 * @param {Function} onError
		 * @param officeResultDeferred
		 *  */
		function getContentControlByTitle(title, syncCallback, onError, officeResultDeferred) {
			// Run a batch operation against the Word object model.
			Word.run(function(context) {

				var contentControls = context.document.contentControls.getByTitle(title);
				context.load(contentControls, 'text, title, tag');
				return context.sync().then(function(asyncResult) {
					syncCallback.call(this, contentControls, context);
				});
			}).catch(function(error) {
				resolveError(error, onError, officeResultDeferred);
			});
		}

		function getDisplayTagToken(displayTag) {
			return "%%" + displayTag + "%%";
		}

		return {
			isOfficeEnabled: function() {
				return Office ? true : false;
			},
			getDownloadUrl: function(onError) {
				try {
					return Office.context.document.url
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}
			},
			getDocumentUrl: function(onError) {
				try {
					var result;
					var url = Office.context.document.url;
					if (url.indexOf(".sharepoint.com") > -1) {
						// is a sharepoint
						var baseUrl = url.substr(0, url.indexOf("download.aspx?"))
						var token = url.substr(url.indexOf("%7B"), url.length)
						//var token = wrappedToken.replace("%7B","").replace("%7D","");
						result = baseUrl + "WopiFrame.aspx?sourcedoc=" + token;
					} else { // one drive
						result = url.replace("download.aspx", "edit.aspx");
					}
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}
				return result;
			},
			getSelectedData: function(type, onError) {

				var officeResultDeferred = $q.defer();
				
				if (!type) type = Office.CoercionType.Text;
				
				try {
					Office.context.document.getSelectedDataAsync(type, getSelectedDataCallback);

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
			setSelectedData: function(text, type, onError) {

				var officeResultDeferred = $q.defer();
				
				if (!type) type = Office.CoercionType.Text;

				try {
					Office.context.document.setSelectedDataAsync(text, { coercionType: type }, setSelectedDataCallback);
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				function setSelectedDataCallback(result) {
					var returnValue;
					if (result.status === Office.AsyncResultStatus.Succeeded) {
						returnValue = true;
					} else {
						resolveError(result.error, onError, officeResultDeferred)
					}
					officeResultDeferred.resolve(returnValue);
				}

				return officeResultDeferred.promise;
			},
			getBodyOoxml: function(onError) {
				/**
				 * Returns the document body as a text
				 *  */
				var officeResultDeferred = $q.defer();
				Word.run(function(ctx) {
					// Create a proxy object for the document body.
					var body = ctx.document.body;
					var result = body.getOoxml();

					return ctx.sync().then(function() {
						$log.debug(result.value)
						officeResultDeferred.resolve(result.value);
					});

				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred);
				});

				return officeResultDeferred.promise;
			},
			getBodyHtml: function(onError) {
				/**
				 * Returns the document body as a text
				 *  */
				var officeResultDeferred = $q.defer();
				Word.run(function(ctx) {
					// Create a proxy object for the document body.
					var body = ctx.document.body;
					var result = body.getHtml();

					return ctx.sync().then(function() {
						$log.debug(result.value)
						officeResultDeferred.resolve(result.value);
					});
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred);
				});

				return officeResultDeferred.promise;
			},
			getBodyText: function(loadOptions, onError) {
				/**
				 * Returns the document body as a text
				 *  */
				var officeResultDeferred = $q.defer();

				// default option
				// TODO should allow or options or return just text ?
				if (!loadOptions) {
					loadOptions = 'text';
				}

				Word.run(function(ctx) {
					// Create a proxy object for the document body.
					var body = ctx.document.body;
					ctx.load(body, loadOptions);
					officeResultDeferred.resolve(body);
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred);
				});

				return officeResultDeferred.promise;
			},
			insertTextToBody: function(text, insertLocation, onError) {
				var officeResultDeferred = $q.defer();

				if (!insertLocation) {
					insertLocation = Word.InsertLocation.end;
				}

				Word.run(insertText).catch(function(e) {
					resolveError(e, onError, officeResultDeferred);
				});

				function insertText(ctx) {
					// Create a proxy object for the document body.
					var body = ctx.document.body;
					ctx.load(body, 'text');
					ctx.document.body.insertText(text, insertLocation);
					return ctx.sync().then(function() {
							officeResultDeferred.resolve(true);
						}, function(e) {
							resolveError(e, onError, officeResultDeferred);
						});
				}

				return officeResultDeferred.promise;
			},
			insertHtmlToBody: function(htmlText, insertLocation, onError) {
				var officeResultDeferred = $q.defer();

				if (!insertLocation) {
					insertLocation = Word.InsertLocation.end;
				}

				Word.run(insertHtml).catch(function(e) {
					resolveError(e, onError, officeResultDeferred);
				});

				function insertHtml(ctx) {
					// Create a proxy object for the document body.
					ctx.document.body.insertHtml(htmlText, insertLocation);
					return ctx.sync().then(function() { // success
							officeResultDeferred.resolve(true);
						}, function(e) { // error
							resolveError(e, onError, officeResultDeferred);
						});
				}

				return officeResultDeferred.promise;
			},
			insertOoxmlToBody: function(ooxmlText, insertLocation, onError) {
				var officeResultDeferred = $q.defer();

				if (!insertLocation) {
					insertLocation = Word.InsertLocation.end;
				}

				Word.run(insertOoxml).catch(function(e) {
					resolveError(e, onError, officeResultDeferred);
				});

				function insertOoxml(ctx) {
					// Create a proxy object for the document body.
					ctx.document.body.insertOoxml(ooxmlText, insertLocation);
					return ctx.sync().then(function() {
							officeResultDeferred.resolve(true);
						}, function(e) {
							resolveError(e, onError, officeResultDeferred);
						});
				}

				return officeResultDeferred.promise;
			},
			selectBody: function(selectionMode, onError) {
				var officeResultDeferred = $q.defer();

				Word.run(function(ctx) {
					var body = ctx.document.body;
					body.select(selectionMode);
					return ctx.sync().then(function() {
						officeResultDeferred.resolve(true);
					});
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				});

				return officeResultDeferred.promise;
			}
		}
	}]).run(function($rootScope, $services) {

	//____logProvider.debugEnabled(true);

	//	requirejs.config({
	//		//By default load any module IDs from js/lib
	//		baseUrl: '/'
	//	});
	//
	//	requirejs(["office365/lib/office.debug.js", "office365/lib/MicrosoftAjax.debug.js"], function(util) {
	//			$log.debug("officejs loaded");
	//		});

});