const Express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const axios = require("axios");

const App = Express();
App.use(bodyParser.urlencoded({ extended: true }));
App.use(Express.json());
App.use(cors());

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
  axios.get(`https://graph.microsoft.com/v1.0/users/{${req.query.userID}}`,{
    headers: {
      'Authorization': `Bearer ${req.headers.authorization}`
        }
    })
    .then((response) => {
      console.log(response.data);
    })
    .catch((error) => {
      console.log(error);
    });

  res.send("Hello world!");
});

const port = process.env.PORT || 8000;
App.listen(port, () => {
  console.log(`Server started on port ${port}`);
});
