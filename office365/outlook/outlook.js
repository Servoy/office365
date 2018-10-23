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
			getSelectedMessageType: function(options, callback, onError){
				
				var officeResultDeferred = $q.defer();
				
				Outlook.run(function(ctx){
					
					var message = ctx.mailbox.item;
					
					var messageType = message.body.getTypeAsync(options, callback)
					
					return ctx.sync().then(function(){
						officeResultDeferred.resolve(messageType);
					});
				}).catch(function(e){
					resolveError(e,onError,officeResultDeferred)
				});
				
				return officeResultDeferred.promise;
			},
		
			getSelectedMessageBody: function(coercionType, options, callback, onError){
				
				var officeResultDeferred = $q.defer();
				Outlook.run(function(ctx){
					
					coercionType = typeof coercionType === 'undefined' ? Office.CoercionType.Text : coercionType;
					
					var message = ctx.mailbox.item;
					
					var bodyText = message.body.getAsync(coercionType, options, callback)
					
					return ctx.sync().then(function(){
						officeResultDeferred.resolve(message);
					});
				}).catch(function(e){
					resolveError(e,onError,officeResultDeferred)
				});
				
				return officeResultDeferred.promise;					
			},
			
			getSelectedMessageSubject:function(){
				var officeResultDeferred = $q.defer();
				Outlook.run(function(ctx){
					
					coercionType = typeof coercionType === 'undefined' ? Office.CoercionType.Text : coercionType;
					
					var message = ctx.mailbox.item;
					
					var messageSubject = ctx.mailbox.item
					
					return ctx.sync().then(function(){
						officeResultDeferred.resolve(messageSubject);
					});
				}).catch(function(e){
					resolveError(e,onError,officeResultDeferred)
				});
				
				return officeResultDeferred.promise;
			},
			
			setSubject: function(){
				
			},
			
			setBodyText: function(data, options, callback, onError){
				
				var officeResultDeferred = $q.defer();
				Outlook.run(function(ctx) {
					
					var messageBody = ctx.mailbox.item.body;
					
					messageBody.setAsync(data, options, callback);
					
					return ctx.sync().then(function(){
						officeResultDeferred.resolve(true);
					})
				}).catch(function(e){
					resolveError(e,onError,officeResultDeferred)
				});
			}
		}
	}]).run(function($rootScope, $services) {
})