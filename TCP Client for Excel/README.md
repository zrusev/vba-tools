TCP Client for Excel.

This is a sample chat client for Excel which uses Windows forms and WebSock library to connect itself to a server and send data.

As it is single thread execution the server should receive a message and stored it.

The next time another client is connected the server sends the whole correspondence so far to the client.

The client receives the whole set of messages only at the moment of sending data.