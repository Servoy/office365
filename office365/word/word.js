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

		// does it make sense ?
		if (Word) {
			var INSERT_LOCATION = {
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
			getSelectedData: function(onError) {

				var officeResultDeferred = $q.defer();
				try {
					Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, getSelectedDataCallback);

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
					//var myPar = parent.ctx.document.body.insertOoxml('<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ><w:body><w:p><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD  my-first-field  \* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="00211A63"><w:rPr><w:noProof/></w:rPr><w:t>«my-first-field»</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>', 'End');

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
			insertContentControl: function(ops, onError) {

				var officeResultDeferred = $q.defer();

				// Run a batch operation against the Word object model.
				Word.run(function(context) {

					// Create a proxy range object for the current selection.
					var range = context.document.getSelection();

					// Queue a commmand to create the content control.
					var contentControl = range.insertContentControl();
					if (ops) {
						if (ops.tag) contentControl.tag = ops.tag;
						if (ops.title) contentControl.title = ops.title;
						if (ops.cannotDelete) contentControl.cannotDelete = ops.cannotDelete;
						if (ops.cannotEdit) contentControl.cannotEdit = ops.cannotEdit;
						if (ops.removeWhenEdited) contentControl.removeWhenEdited = ops.removeWhenEdited;
						if (ops.placeholderText) contentControl.placeholderText = ops.placeholderText;
						if (ops.color) contentControl.color = ops.color;
						if (ops.style) contentControl.style = ops.style;
					}

					// Queue a command to load the id property for the content control you created.
					context.load(contentControl, 'id');

					// Synchronize the document state by executing the queued commands,
					// and return a promise to indicate task completion.
					return context.sync().then(function() {
							var id = contentControl.id;
							if (typeof (id) === "number") {
								id = "" + id;
							}
							officeResultDeferred.resolve(id);
						}, function(e) {
							resolveError(e, onError, officeResultDeferred)
						});
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				});

				return officeResultDeferred.promise;
			},

			selectBody: function(selectionMode, onError) {
				var officeResultDeferred = $q.defer();

				Word.run(function(ctx) {
					// Create a proxy object for the document body.
					var body = ctx.document.body;
					body.select(selectionMode);
					return ctx.sync().then(function() {
						officeResultDeferred.resolve(true);
					});
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				})

				return officeResultDeferred.promise;
			},
			searchBody: function(searchText, searchOptions, onError) {
				var officeResultDeferred = $q.defer();

				Word.run(function(ctx) {
					// Create a proxy object for the document body.
					var body = ctx.document.body;
					var searchResults = body.search(searchText, searchOptions);
					// TODO what else to include in the search ?
					if (scope.model.searchResults) {
						ctx.trackedObjects.remove(scope.model.searchResults);
					}
					ctx.load(searchResults, 'text, id, title, font/size');
					return ctx.sync().then(function() {
						// store the searchResults in cache
						$log.debug(searchResults);
						ctx.trackedObjects.add(searchResults);
						scope.model.searchResults = searchResults.items;
						officeResultDeferred.resolve(searchResults.items);
					});
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				})

				return officeResultDeferred.promise;
			},
			goToContentControl: function(title, selectionMode, onError) {
				var officeResultDeferred = $q.defer();

				getContentControlByTitle(title, syncCallback, onError, officeResultDeferred)

				function syncCallback(contentControls) {
					if (contentControls.items.length === 0) {
						officeResultDeferred.resolve(false);
					} else {
						var nextItem = contentControls.items[0].select(selectionMode);
						officeResultDeferred.resolve(true);
					}
				}

				return officeResultDeferred.promise;

			},
			getAllContentControls: function(onError) {
				var officeResultDeferred = $q.defer();

				// Run a batch operation against the Word object model.
				Word.run(function(context) {

					var contentControls = context.document.contentControls;
					context.load(contentControls, 'text, title, tag');
					return context.sync().then(function() {
						officeResultDeferred.resolve(contentControls.items);
					});
				}).catch(function(error) {
					resolveError(error, onError, officeResultDeferred);
				});

				return officeResultDeferred.promise;
			},
			setContentControlText: function(title, text, onError) {
				/***
				 * Find all content control having title
				 * Replace the text for each item found
				 * Using binding to replace text is prefered to plain insertText since using binding is possible to mantain the design text style.
				 * insertText will reset the text style to normal instead.
				 * Since there may be several item with the same namesSpace (title),
				 * - select the item
				 * - add binding from selection
				 * - set binding data
				 * - release time
				 *
				 * */

				var officeResultDeferred = $q.defer();

				//getContentControlByTitle(title, syncCallback, onError, officeResultDeferred);

				//				// go to all of them
				//				function syncCallback(contentControls, context) {
				//					if (contentControls.items.length === 0) {
				//						officeResultDeferred.resolve(false);
				//					} else {
				//						for (var i = 0; i < contentControls.items.length; i++) {
				//							contentControls.items[i].insertText(text, "replace");
				//						}
				//						officeResultDeferred.resolve(true);
				//
				//						context.sync().then(function() {
				//							officeResultDeferred.resolve(true);
				//						}).catch(function(e) {
				//							resolveError(e, onError, officeResultDeferred)
				//						})
				//					}
				//				}

				// Run a batch operation against the Word object model.
				Word.run(function(context) {

					var contentControls = context.document.contentControls.getByTitle(title);
					context.load(contentControls, 'text, title, tag');
					// sync
					return context.sync().then(function(asyncResult) {
						// wait until process is complete
						return syncCallback(contentControls, context).then(function(result) {
							officeResultDeferred.resolve(result);
						}).catch(function(error) {
							resolveError(error, onError, officeResultDeferred);
						});
					});
				}).catch(function(error) {
					resolveError(error, onError, officeResultDeferred);
				});

				// 0 search all content control by title
				function syncCallback(contentControls, context) {
					
					// use promise to force Run method to wait until process is complete
					return new Promise(function(resolveSync, rejectSync) {
						if (contentControls.items.length === 0) {
							resolveSync(0);
						} else {
							var i = 0;
							iterate();

							// iterate over each item of the contentControl
							function iterate() {
								var item = contentControls.items[i];
								if (i < contentControls.items.length) {		// until has item to process
									$log.debug('setContentControlTex start ' + item.id);

									i++;
									processItem(item, context).then(function() {
										iterate();
									}).catch(function(error) {
										resolveError(error, onError, officeResultDeferred);
									});

								} else {	// process is complete, sync to push all outstanding changes to document and resolve
									return context.sync().then(function() {
										resolveSync(i);
									});
								}
							}
						}
					});
				}

				function processItem(item, context) {

					/* Use promise to make sure that value are processed sequentially */
					return new Promise(function(resolve, reject) {

						// 1 select item
						item.select('select')
						return item.context.sync().then(function() {
							$log.debug('setContentControlTex selected ' + item.id)
							var itemId = item.id;
							if (typeof (itemId) === "number") {
								itemId = itemId + "";
							}
							// add binding
							addBindingFromSelection(itemId);
						}).catch(function(e) {
							resolveError(e, onError, officeResultDeferred)
						})

						// 2 add binding
						function addBindingFromNameSpace(id) {
							Office.context.document.bindings.addFromNamedItemAsync(title, Office.CoercionType.Text, { "id": id, asyncContext: id }, replaceBindingsText);
						}

						function addBindingFromSelection(id) {
							Office.context.document.bindings.addFromSelectionAsync(Office.CoercionType.Text, { "id": id, asyncContext: id }, replaceBindingsText);
						}

						// 3 replace text
						function replaceBindingsText(asyncRes) {

							$log.debug('setContentControlTex setSelectedBinding ' + item.id)

							var id = asyncRes.asyncContext;

							// 3 set binding to the selected text
							if (asyncRes.status === Office.AsyncResultStatus.Succeeded) {
								Office.select("bindings#" + id, resolveError).setDataAsync(text, setBindingDataCallback);
							} else {
								reject(asyncRes.error);
							}

							// 4 replace the text of the binding callback
							function setBindingDataCallback(result) {
								$log.debug('setContentControlTex replaced ' + item.id)

								if (result.status === Office.AsyncResultStatus.Succeeded) {
									Office.context.document.bindings.releaseByIdAsync(id, releaseByIdCallback);
								} else {
									reject(result.error);
								}
							}

							// 5 release the binding callback
							function releaseByIdCallback(asyncResult) {

								$log.debug('setContentControlTex released ' + item.id);

								$log.debug(asyncResult);
								if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
									// 3 replace the text of the binding
									resolve(true);
								} else {
									reject(asyncResult.error);
								}
							}
						}
					});

				}

				return officeResultDeferred.promise;
			},
			deleteContentControl: function(title, keepContent, onError) {
				var officeResultDeferred = $q.defer();

				getContentControlByTitle(title, syncCallback, onError, officeResultDeferred);

				function syncCallback(contentControls) {
					if (contentControls.items.length === 0) {
						officeResultDeferred.resolve(false);
					} else {
						contentControls.items[0].delete(keepContent);
						officeResultDeferred.resolve(true);
					}
				}

				return officeResultDeferred.promise;
			},
			getSearchResults: function() {
				return scope.model.searchResults;
			},
			selectSearchResult: function(index, selectionMode, onError) {

				// search result not available
				if (!scope.model.searchResults || !scope.model.searchResults[index]) {
					return false;
				}

				var searchResult = scope.model.searchResults[index];

				var officeResultDeferred = $q.defer();

				Word.run(function(ctx) {
					searchResult.contentControls.select(selectionMode);
					return ctx.sync().then(function() {
						// store the searchResults in cache
						officeResultDeferred.resolve(true);
					});
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				})

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