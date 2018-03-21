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

///////////////////////////////////////////////////////////
/*
var bot= new builder.UniversalBot(connector,function(session){
	var msg ='Welcome to reservation.Say dinner or order or test ';
	session.send(msg);
}).set('storage',inMemoryStorage);

bot.dialog('dinnerReservation',[
	function(session){
		session.send('Welcome to ReeServe');
		session.beginDialog('askForDateTime');
	},
	function(session,results){
		session.dialogData.reservationDate=builder.EntityRecognizer.resolveTime([results.response]);
		session.beginDialog('askForPartySize');
	},
	function(session,results){
		session.dialogData.partySize=results.response;
		session.beginDialog('askForReserverName');
	},
	function(session,results){
		session.dialogData.reservationName=results.response;

		session.send(`Reservation confirmed.Reservation details: <br/>Date/Time: ${session.dialogData.reservationDate} <br/>Party Size: ${session.dialogData.partySize} <br/>ReservationName: ${session.dialogData.reservationName}`);
		session.endDialog();
	}
	]).triggerAction({
		matches: /^dinner$/i,
		confirmPrompt:'This will you cancel your previous request.Are you sure?'
	})
	.endConversationAction(
			"endDinner","Ok.Goodbye.",
			{
				matches:/^cancel$|^goodbye$/i,
				confirmPrompt:'This will cancel your dinner order and close!Sure?'
			});;
	

	var dinnerMenu ={
		"Potato Salad - $5.99":{
			Description:"Potato salad",
			Price:5.99
		},
		"Tuna Sandwich - $6.89": {
        Description: "Tuna Sandwich",
        Price: 6.89
    },
    "Clam Chowder - $4.50":{
        Description: "Clam Chowder",
        Price: 4.50
    }
	};

	bot.dialog('orderDinner',[
		function(session){
			session.send('Lets order some dinner');
			builder.Prompts.choice(session,"Dinner Menu: ",dinnerMenu,{listStyle: builder.ListStyle.button});
		},
		function(session,results){
			if(results.response){
				var order=dinnerMenu[results.response.entity];
				var msg = `You ordered: ${order.Description} for a total of $${order.Price}.`;

				session.dialogData.order= order;

				session.send(msg);
				builder.Prompts.text(session,'Where do you live?');
			}
			},
			function(session,results){
				if(results.response){
					session.dialogData.room=results.response;
					var msg=`Thank you.Your order will be delivered to room #${session.dialogData.room}`;
					session.endDialog(msg);
				}
			}
		
		]).triggerAction({
			matches:/^order$/i,
			confirmPrompt:'This will cancel your Order,are you sure'
		})
		.endConversationAction(
			"endOrderDinner","Ok.Goodbye.",
			{
				matches:/^cancel$|^goodbye$/i,
				confirmPrompt:'This will cancel your order and close!Sure?'
			});


		bot.dialog('test',function(session){
			builder.Prompts.attachment(session,'Are you sure');
		
		})
		.triggerAction({
			matches:/^test$/i,
			confirmPrompt:'This will cancel your Order,are you sure'
		})
	
	

	//task automation


*/
//Sub-Dialogs
bot.library(require('./dialogs/reset-password'));

//Validators
bot.library(require('./validators'));

server.post('/api/messages', connector.listen());