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