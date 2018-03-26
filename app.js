var builder = require('botbuilder');
var restify = require('restify');


var connector = new builder.ChatConnector({
	appId: process.env.MicrosoftAppId,
	appPassword: process.env.MicrosoftAppPassword
});
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function(){
	console.log('%s listening to %s ',server.name,server.url);
})
server.post('/api/messages',connector.listen());


var inMemoryStorage = new builder.MemoryBotStorage();

var helpOptions={
	"ColorCode":{
		Contact:"richard@kpmg.co.ke"
	},
	"GreatSoft":{
		Contact:"onditi@kpmg.co.ke"
	},
	"Emails":{
		Contact:"kyalo@kpmg.co.ke"
	},
};

var bot = new builder.UniversalBot(connector, [
    (session) => {
        builder.Prompts.choice(session,
            'Hi am here to help.Choose what is troubling you',
           helpOptions,
            { listStyle: builder.ListStyle.button });
    },
    (session, result) => {
        if (result.response) {
        	
            switch (result.response.entity) {
                case 'GreatSoft':
                    session.beginDialog('greatsoftDialog');
                    break;
                case 'Emails':
                   session.send(`Initializing chat with ${helpOptions[result.response.entity].Contact}`);
                   session.reset();
                    break;
                case 'ColorCode':
                	session.beginDialog('colorcodeDialog',helpOptions[result.response.entity].Contact);
                    break;
                	
            }
        } else {
            session.send(`I am sorry but I didn't understand that. I need you to select one of the options below`);
        }
    },
    (session, result) => {
        if (result.resume) {
            session.send('You identity was not verified and your password cannot be reset');
            session.reset();
        }
    }
]);

const ChangePasswordOption = 'Change Password';
const ResetPasswordOption = 'Reset Password';

bot.dialog('greatsoftDialog', [
    (session) => {
        builder.Prompts.choice(session,
            'What would you like to get help on',
            [ChangePasswordOption, ResetPasswordOption],
            { listStyle: builder.ListStyle.button });
    },
    (session, result) => {
        if (result.response) {
            switch (result.response.entity) {
                case ChangePasswordOption:
                    session.send('This functionality is not yet implemented! Try resetting your password.');
                    session.reset();
                    break;
                case ResetPasswordOption:
                    session.beginDialog('resetPassword:/');
                    break;
            }
        } else {
            session.send(`I am sorry but I didn't understand that. I need you to select one of the options below`);
        }
    },
    (session, result) => {
        if (result.resume) {
            session.send('You identity was not verified and your password cannot be reset');
            session.reset();
        }
    }
]);

bot.dialog('colorcodeDialog', [
	
		function(session,args){
				session.dialogData.senderrorto=args;
				builder.Prompts.text(session,'Please enter the error code recieved or briefly describe your issue?');	
		},
		function(session,results){
					session.dialogData.error=results.response;
					openOutlook(session.dialogData.senderrorto,session.dialogData.error);
					var msg=`Thank you for reporting.Issue sent to ${session.dialogData.senderrorto}  for action`;
					session.endDialog(msg);
				}
			

]);

bot.dialog('helpDialog', [
	
		function(session,args){
				session.dialogData.senderrorto=args;
				builder.Prompts.text(session,'I am here to chat your way out of your technical issues');
				session.reset();
		}
])		
.triggerAction({
			matches:/^help$/i,
			confirmPrompt:'This will interrupt your conversation,are you sure?'
		});

bot.dialog('outlookRead', [
	
		function(session){
				
				builder.Prompts.attachment(session,'Attach outlook object');
				
		},
		function(session,results){
		//read email object
		var url=results.response[0].contentUrl;

	var http = require('http');

//
var emailbody='';
http.get(url, function(res) {
    var data = [];

    res.on('data', function(chunk) {
        data.push(chunk);
    }).on('end', function() {
        //at this point data is an array of Buffers
        //so Buffer.concat() can make us a new Buffer
        //of all of them together
        var buffer = Buffer.concat(data);
        emailbody=readBuffer(buffer);
    });
    
});
//sentiment analysis ALGORITHIMIA
var input = {
  "document": emailbody
};

var Algorithmia=require('algorithmia');
Algorithmia.client("simkW3Zwdt2gz7anbSf62wu7KzS1")
    .algo("nlp/SentimentAnalysis/1.0.4")
    .pipe(input)
    .then(function(response) {
        console.log('Email Sentiment score->: '+response.result[0].sentiment);
    });
}
])		
.triggerAction({
			matches:/^outlook$/i,
			confirmPrompt:'This will interrupt your conversation,are you sure?'
		});

function readBuffer(buffer){
	
  function formatEmail(data) {
    return data.name ? data.name + " [" + data.email + "]" : data.email;
  }

  function parseHeaders(headers) {
    var parsedHeaders = {};
    if (!headers) {
      return parsedHeaders;
    }
    var headerRegEx = /(.*)\: (.*)/g;
    while (m = headerRegEx.exec(headers)) {
      // todo: Pay attention! Header can be presented many times (e.g. Received). Handle it, if needed!
      parsedHeaders[m[1]] = m[2];
    }
    return parsedHeaders;
  }

  function getMsgDate(rawHeaders) {
    // Example for the Date header
    var headers = parseHeaders(rawHeaders);
    if (!headers['Date']){
      return '-';
    }
    return new Date(headers['Date']);
  }


      const MSGReader= require ('./mail_reader/msg.reader.js');

      //require '../msg.reader.js'

          var msgReader = new MSGReader(buffer);
         
          var fileData = msgReader.getFileData();
          console.log('fileData'+fileData);
          if (!fileData.error) {
            console.log('name: '+fileData.senderName );
            console.log(' email: '+ fileData.senderEmail);
            console.log(' sent to: '+ fileData.recipients);

            // $('.msg-example .msg-to').html(jQuery.map(fileData.recipients, function (recipient, i) {
            //   return formatEmail(recipient);
            // }).join('<br/>'));
           console.log('Email Date= '+getMsgDate(fileData.headers));
            console.log('Email Subject= '+fileData.subject);
            console.log('Email Body'+fileData.body);
            console.log('attachments'+Object.keys(fileData.attachments));
            // $('.msg-example .msg-attachment').html(jQuery.map(fileData.attachments, function (attachment, i) {
            //   return attachment.fileName + ' [' + attachment.contentLength + 'bytes]' +
            //       (attachment.pidContentId ? '; ID = ' + attachment.pidContentId : '');
            // }).join('<br/>'));
            // $('.msg-info').show();

            // Use msgReader.getAttachment to access attachment content ...
            // msgReader.getAttachment(0) or msgReader.getAttachment(fileData.attachments[0])
          } else {
           console.log('Parsed message has an error');
          }
        
        
     
   return  fileData.body;
}
  



function openOutlook(sendto,error){

		var childProcess = require('child_process');

    	childProcess.spawn(
        "powershell.exe",
        ['$mail = (New-Object -comObject Outlook.Application).CreateItem(0);$mail.Subject = "Color Code Error";$mail.Body="I have error '+error+' on Color Code.Kindly Assist";$mail.To ="'+sendto+'"; $mail.Send();']
    );

}

//Sub-Dialogs
bot.library(require('./dialogs/reset-password'));

//Validators
bot.library(require('./validators'));

server.post('/api/messages', connector.listen());

//outlook read metadata handlers



  function formatEmail(data) {
    return data.name ? data.name + " [" + data.email + "]" : data.email;
  }
  function parseHeaders(headers) {
    var parsedHeaders = {};
    if (!headers) {
      return parsedHeaders;
    }
    var headerRegEx = /(.*)\: (.*)/g;
    while (m = headerRegEx.exec(headers)) {
      // todo: Pay attention! Header can be presented many times (e.g. Received). Handle it, if needed!
      parsedHeaders[m[1]] = m[2];
    }
    return parsedHeaders;
  }
  function getMsgDate(rawHeaders) {
    // Example for the Date header
    var headers = parseHeaders(rawHeaders);
    if (!headers['Date']){
      return '-';
    }
    return new Date(headers['Date']);
  }

