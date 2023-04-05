const MicrosoftGraph = require("@microsoft/microsoft-graph-client");
const graphAuth = require("@azure/msal-node");
const express = require("express");
const bodyParser = require('body-parser');
const session = require('express-session');
const app = express();
const port = process.env.PORT || 3000;

// Define the required environment variables
const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;

// Get an access token for the Microsoft Graph API
async function getAccessToken() {
  const cca = new graphAuth.ConfidentialClientApplication({
    auth: {
        clientId: clientId,
        clientSecret: clientSecret,
        authority: `https://login.microsoftonline.com/${tenantId}`
    }
  });
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"]
  });
  return result.accessToken;
}

// Create an Express.js app
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(session({
  secret: 'secret',
  resave: false,
  saveUninitialized: false
}));

// Serve a form to allow users to opt-in to the tool
app.get("/", (req, res) => {
  res.send(`
    <form method="POST" action="/opt-in">
      <label for="email">Email address:</label>
      <input type="email" id="email" name="email" required>
      <button type="submit">Opt-in</button>
    </form>
  `);
});

// Handle POST requests to opt-in to the tool
app.post("/opt-in", (req, res) => {
  const email = req.body.email;

  // Get an access token for the Microsoft Graph API
  getAccessToken().then((accessToken) => {

    // Create a Microsoft Graph client with the access token
    const client = MicrosoftGraph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    // Create an extension property for the user to indicate that they have opted-in
    client.api(`/users/${email}/extensions`)
      .post({
        "@odata.type": "microsoft.graph.openTypeExtension",
        "extensionName": "extension_fe7bfb23fca34e7fbae56894d9791c83_MurraysPartnershipAlert",
        "extension_fe7bfb23fca34e7fbae56894d9791c83_optedIn": true
      }, (err, result) => {
        if (err) {
          console.error(err);
          res.status(500).send("Error opting in. Please try again.");
        } else {
          res.send(`
            <p>You have successfully opted-in to the tool</p>
            <p><a href="/alert-users">Send upcoming partnership alert</a></p>
          `);
        }
      });

  }).catch((err) => {
    console.error(err);
    res.status(500).send("Error retrieving access token. Please try again.");
  });
});

// Serve a form to allow authorized users to send an email to users who have opted-in
app.get("/alert-users", (req, res) => {
  const email = req.query.email;

  // Get an access token for the Microsoft Graph API
  getAccessToken().then((accessToken) => {

    // Create a Microsoft Graph client with the access token
    const client = MicrosoftGraph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    // Check if the user has opted-in to the tool
    client.api(`/users/${email}`)
      .select("id")
      .expand("extensions")
      .get((err, result) => {
        if (err) {
          console.error(err);
          res.status(500).send("Error retrieving user information. Please try again.");
        } else {
          const optedIn = result.extensions.some((extension) => {
            return extension.extensionName === "extension_fe7bfb23fca34e7fbae56894d9791c83_MurraysPartnershipAlert" && extension.extension_fe7bfb23fca34e7fbae56894d9791c83_MurraysPartnershipAlert_optedIn;
          });

          if (optedIn) {
            // Serve the email form
            res.send(`
              <form method="POST" action="/email-users">
                <input type="hidden" name="email" value="${email}">
                <textarea name="message" required></textarea>
                <br>
                <button type="submit">Send partnership alert</button>
              </form>
            `);
          } else {
            // User has not opted-in to the tool
            res.status(403).send("Access denied.");
          }
        }
      });

  }).catch((err) => {
    console.error(err);
    res.status(500).send("Error retrieving access token. Please try again.");
  });
});

// Handle POST requests to send an email to users who have opted-in
app.post("/email-users", (req, res) => {
  const email = req.body.email;
  const message = req.body.message;

  // Get an access token for the Microsoft Graph API
  getAccessToken().then((accessToken) => {

    // Create a Microsoft Graph client with the access token
    const client = MicrosoftGraph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    // Find all users who have opted-in to the tool and have a calendar event with the specified email address in the last three months
    client.api(`/users`)
      .select("id,displayName,mail")
      .expand("extensions")
      .filter(`extensions/any(extension: extension/id eq 'extension_fe7bfb23fca34e7fbae56894d9791c83_MurraysPartnershipAlert' and extension/fe7bfb23fca34e7fbae56894d9791c83_MurraysPartnershipAlert_optedIn eq true)`)
      .expand("calendarView($filter=start/dateTime ge '${new Date(Date.now() - 777600000).toISOString()}')")
      .get((err, result) => {
        if (err) {
          console.error(err);
          res.status(500).send("Error retrieving user information. Please try again.");
        } else {
          const users = result.value.filter((user) => {
            return user.calendarView.some((event) => {
              return event.attendees.some((attendee) => {
                return attendee.emailAddress.address.toLowerCase() === email.toLowerCase();
              });
            });
          });

          // Send an email to each user who has opted-in and has a matching calendar event
          users.forEach((user) => {
            const email = {
              subject: "Upcoming Partnership Alert",
              toRecipients: [{
                emailAddress: {
                  address: user.mail
                }
              }],
              body: {
                content: message,
                contentType: "text"
              }
            };
            client.api("/me/sendMail")
              .post({ message: email }, (err, res) => {
                if (err) {
                  console.error(err);
                } else {
                  console.log(`Email sent to ${user.displayName} (${user.mail}).`);
                }
              });
          });

          res.send(`
            <p>Emails have been sent to all users who have opted-in and have a matching calendar event.</p>
          `);
        }
      });

  }).catch((err) => {
    console.error(err);
    res.status(500).send("Error retrieving access token. Please try again.");
  });
});

// Start the server
app.listen(port, () => {
  console.log(`Server started on port ${port}.`);
});
