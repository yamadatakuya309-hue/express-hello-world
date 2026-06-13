const express = require('express');
const line = require('@line/bot-sdk');

const app = express();

const config = {
  channelAccessToken: 'a9Gm8O+G17f9rIcBAiZ1cD5O1AR2jyPrDJYZfxlYbn0S9+igtU4GIyWyB2hcwSwXkflxCKifxFZGd0YuhTfqfHqxnZWgyG91K6u0X1qhPebtKfsuIrfRoL058XvAI+FLA1PwrNUwjCCm8WPmN67hwAdB04t89/1O/w1cDnyilFU=',
channelSecret: '14bd181a88c13f4e28d0b28f6626c682'
};

const client = new line.Client(config);

app.post('/webhook', line.middleware(config), (req, res) => {
  Promise.all(req.body.events.map(handleEvent))
    .then((result) => res.json(result))
    .catch((err) => {
      console.error(err);
      res.status(500).end();
    });
});

function handleEvent(event) {
  if (event.type !== 'message' || event.message.type !== 'text') {
    return Promise.resolve(null);
  }
  const userMessage = event.message.text;
  return client.replyMessage(event.replyToken, {
    type: 'text',
    text: '受け取りました：' + userMessage
  });
}

const port = process.env.PORT || 10000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
