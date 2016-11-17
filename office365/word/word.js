angular.module('office365Word',['servoy'])
.factory("office365Word", ['$services', '$window', '$log', function($services, $window, $log) 
{
	var scope = $services.getServiceScope('office365Word');
	var Office = parent.Office;
	
	return {
		getDownloadUrl: function(onError) {
			try {
				return Office.context.document.url
			} catch (e) {
				$log.error('office365Word: error getting document URL: ' + e.message);
				if (onError) {
					$window.executeInlineScript(onError.formname, onError.script, [e.message]);
				}
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
				} else {	// one drive
					result = url.replace("download.aspx", "edit.aspx");
				}
			} catch (e) {
				$log.error('office365Word: error getting document URL: ' + e.message);
				if (onError) {
					$window.executeInlineScript(onError.formname, onError.script, [e.message]);
				}
			}
			return result;
		},
		getSelectedData: function(callbackMethod, onError) {
			
			Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, getSelectedDataCallback);
			
			function getSelectedDataCallback (result) {
				if (result.status === Office.AsyncResultStatus.Succeeded) {
					$window.executeInlineScript(callbackMethod.formname, callbackMethod.script, [result.value]);
				} else {
					$log.error('office365Word: ', result.error.message);
					$window.executeInlineScript(onError.formname, onError.script, [result.error.message]);
					return null;
				}
			}
			return;
		},
		setSelectedData: function(text, type, callbackMethod, onError) {
			
				try {
					Office.context.document.setSelectedDataAsync(text, {coercionType: type}, setSelectedDataCallback);						
				} catch (e) {
					$log.error('office365Word: error setting selected data: ' + e.message);
					$window.executeInlineScript(onError.formname, onError.script, [e.message]);
				}
				
				function setSelectedDataCallback (result) {
					if (result.status === Office.AsyncResultStatus.Succeeded) {
						$window.executeInlineScript(callbackMethod.formname, callbackMethod.script, [result.value]);
						return result;
					} else {
						$log.error('office365Word: ', result.error.message);
						$window.executeInlineScript(onError.formname, onError.script, [result.error.message]);
						return null;
					}
				}
		},
		insertText: function(text, onSuccess, onError) {
			// call the servoy method from callback success
				try {
			        parent.Word.run(function (ctx) {
			            // Queue a command to add a paragraph at the end of the document. Set the style on the paragraph.
			            var myPar = parent.ctx.document.body.insertText(text, 'End');
			        });
				} catch (e) {
					$log.error('office365Word: error getting selected data: ' + e.message);
				}
				
				function onSuccessCallback (result) {
//					if (result.status === Office.AsyncResultStatus.Succeeded) {
//	 					Bridge.iFrame.contentWindow.postMessage({ method: successCallback, params: result.value }, '*');
//						return result;
//					} else {
//						console.log('Error:', result.error.message);
//						return null;
//					}
				}
		},
		insertMergeField: function(text, onSuccess, onError) {
			
			// call the servoy method from callback success

				try {
			        parent.Word.run(function (ctx) {
			            // Queue a command to add a paragraph at the end of the document. Set the style on the paragraph.
			            //var myPar = ctx.document.body.insertOoxml(wordtext, 'End');
			            var myPar = parent.ctx.document.body.insertOoxml('<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ><w:body><w:p><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD  my-first-field  \* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="00211A63"><w:rPr><w:noProof/></w:rPr><w:t>«my-first-field»</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>', 'End');
			            
			            return ctx.sync()
			            .then(function () {
			            	// TODO enable debug
			                console.log("insert ooxml success")
			            })
			            .catch(function (error) {
			                console.log("insert ooxml error")
							$log.error(error)
			            })
			        });
				} catch (e) {
					$log.error('office365Word: error getting selected data: ' + e.message);
				}
				
				function onSuccessCallback (result) {
//					if (result.status === Office.AsyncResultStatus.Succeeded) {
//	 					Bridge.iFrame.contentWindow.postMessage({ method: successCallback, params: result.value }, '*');
//						return result;
//					} else {
//						console.log('Error:', result.error.message);
//						return null;
//					}
				}
		},
		writeDocx: function(xml, onSuccess, onError) {
			

				try {
			        parent.Word.run(function (ctx) {
			            // Queue a command to add a paragraph at the end of the document. Set the style on the paragraph.
			            var myPar = parent.ctx.document.body.insertOoxml(xml, 'Replace');
			            
			            return ctx.sync()
			            .then(function () {
			                console.log("write xml success")
			            })
			            .catch(function (error) {
			                console.log("write xml error")
			            })
			            
			        });
				} catch (e) {
					window.alert('error getting selected data: ' + e.message);
				}
				
				function onSuccessCallback (result) {
//					if (result.status === Office.AsyncResultStatus.Succeeded) {
//	 					Bridge.iFrame.contentWindow.postMessage({ method: successCallback, params: result.value }, '*');
//						return result;
//					} else {
//						console.log('Error:', result.error.message);
//						return null;
//					}
				}
		}
	}
}]).run(function($rootScope, $services) {
	
})