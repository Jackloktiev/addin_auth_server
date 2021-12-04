const Express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const axios = require("axios");
const mongoose = require("mongoose"); //to connect to the Mongo Atlas

const App = Express();
App.use(bodyParser.urlencoded({ extended: true }));
App.use(Express.json());
App.use(cors());

//Connect to MongoDB
mongoose
  .connect(
    "mongodb+srv://jack_olsenconsulting:Winter2022!@cluster0.xtifk.mongodb.net/authorization_db?retryWrites=true&w=majority",
    { useNewUrlParser: true, useUnifiedTopology: true }
  )
  .then((res) => console.log("Connected to DB"))
  .catch((err) => console.log(err));

// Create user schema
const userShema = new mongoose.Schema({
  userDomain: String, // i.e. "@olsenconsulting.ca"
  excelTemplateSAS: Array, // i.e. ["https://olsenconsultingaddn.blob.core.windows.net/sharedfiles/AR%20Reconciliation%20Workbook%20V2.xlsm?sp=r&st=2021-11-21T17:33:22Z&se=2023-01-01T01:33:22Z&spr=https&sv=2020-08-04&sr=b&sig=0xMdA8NfPxON4ETOLJ6eoeYiCiA9s57sjOhXwkolCv8%3D", ""]
});

const User = mongoose.model("User", userShema);

App.get("/hello", function (req, res) {
  res.send("Hello world!");
});

App.get("/gettemplate", function (req, res) {
  console.log(req.headers.authorization);
  console.log(req.query.file);
  console.log(req.query.userID);
  // Get user email from MS graph
  //   send get request to https://graph.microsoft.com/v1.0/users/{req.query.userID}
  // with authorization header = 'Bearer {token}'

  axios
    .get(`https://graph.microsoft.com/v1.0/users/{${req.query.userID}}`, {
      headers: {
        Authorization: `Bearer ${req.headers.authorization}`,
      },
    })
    .then((response) => {
      // When we get back data from MS Graph take user domain and look it up in Mongo DB
      console.log(response.data);
      const domain = response.data.mail.slice(response.data.mail.indexOf("@"));

      User.findOne({ userDomain: domain }, function (error, foundUser) {
        if (error) {
          console.log(error);
        } else {
          console.log("user found in the DB");
          // if user is not found in the db the foundUser==null
          if (foundUser) {
            console.log(foundUser.excelTemplateSAS);
            //Send a list of SAS URLs back to the add-in to build a list of files available to the user
            res.send(foundUser.excelTemplateSAS);
          } else {
            res.status(600);
            res.send("User not authorized");
          }
        }
      });

      // Code to add new user to the DB
      // const newUser = new User({
      //   userDomain: "@olsenconsulting.ca",
      //   excelTemplateSAS: [
      //     ["AR_recon", "https://olsenconsultingaddn.blob.core.windows.net/sharedfiles/AR%20Reconciliation%20Workbook%20V2.xlsm?sp=r&st=2021-11-21T17:33:22Z&se=2023-01-01T01:33:22Z&spr=https&sv=2020-08-04&sr=b&sig=0xMdA8NfPxON4ETOLJ6eoeYiCiA9s57sjOhXwkolCv8%3D"],
      //     ["AP_recon","https://olsenconsultingaddn.blob.core.windows.net/sharedfiles/AP%20Reconciliation%20Workbook%20V6.xlsm?sp=r&st=2021-11-21T17:52:32Z&se=2023-01-01T01:52:32Z&spr=https&sv=2020-08-04&sr=b&sig=vrk%2FRNNfoC29qYQ98rIhkl%2FbspKy5eZB2BitinRpQ%2BA%3D"],
      //     ["JC_detail","https://olsenconsultingaddn.blob.core.windows.net/sharedfiles/JC%20Detail%20for%20T%26M%20Billing.xlsx?sp=r&st=2021-11-21T17:54:02Z&se=2023-01-01T01:54:02Z&spr=https&sv=2020-08-04&sr=b&sig=2TN3%2FZW9smnTRxuDySh7fEJo2eGD4iSNtCuY6NpZcDM%3D"],
      //     ["Prod_tracker","https://olsenconsultingaddn.blob.core.windows.net/sharedfiles/Productivity%20Tracker%20-%20for%20Sunny.xlsm?sp=r&st=2021-11-21T17:53:27Z&se=2023-01-01T01:53:27Z&spr=https&sv=2020-08-04&sr=b&sig=Nq7wGlQDIXoaI078kMWKLn%2BiaIduILeqb4Ma0VPm9sE%3D"]
      //   ]

      // });

      // newUser.save(function(error){
      //   if(error){
      //     console.log(error)
      //   }else{
      //     console.log("User saved")
      //   }
      // })
    })
    .catch((error) => {
      console.log(error);
    });
});

const port = process.env.PORT || 8000;
App.listen(port, () => {
  console.log(`Server started on port ${port}`);
});
