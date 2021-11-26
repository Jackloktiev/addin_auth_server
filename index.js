const Express = require("express");
const bodyParser = require("body-parser");
const cors = require('cors');

const App = Express();
App.use(bodyParser.urlencoded({extended:true}));
App.use(Express.json());
App.use(cors());

App.get("/hello",function(req,res){
    res.send('Hello world!');
});

const port = process.env.PORT || 8000;
App.listen(port, () => {
    console.log(`Server started on port ${port}`);
});