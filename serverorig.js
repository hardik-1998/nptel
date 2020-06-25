const express = require("express");
const excel = require("exceljs");
var mysql = require('mysql');
const body = require("body-parser");
// var connection = mysql.createConnection({
//     host: 'localhost',
//     user: 'root',
//     password: '',
//     database: 'nptel'
// })
var connection = mysql.createPool({
    host:'us-cdbr-east-06.cleardb.net',
    port:3306,
    user:'b3ef2efda41813',
    password:'44f4c812',
    database:'heroku_4a6cdaac807847d'
})

// connection.connect(function (error) {

//     if (!!error) {
//         console.log('Error');
//     } else {

//         console.log('Connected');
//     }
// });
connection.on('error', function (err) {
    console.log("caught this error" + err.toString());
});
var app = express();
app.set('port',(process.env.PORT || 8081));
app.use(express.static("public"));
app.use(body.urlencoded({ extended: false }))

app.post("/formlogin", function (req, res) {
    var name = req.body.first_name;

    var dbexcel = {
        name: req.body.first_name,
        lastname: req.body.last_name,
        birthday: req.body.birthday,
        roll: req.body.roll,
        caste: req.body.caste,
        gender: req.body.gender,
        email: req.body.email,
        phone: req.body.phone,
        year: req.body.year,
        branch: req.body.branch,
        position: req.body.position,
        coursename: req.body.coursename
    }
    var nsql = "SELECT * FROM ?? WHERE NOT EXISTS (SELECT NULL FROM ?? WHERE ??=? )";

    var nrecords = ['registration', 'Email', dbexcel.email];
    connection.query(nsql, nrecords, function (error, rows, fields) {
        if (error) {
         var sql = "INSERT INTO registration VALUES ?";
            var records = [dbexcel.name, dbexcel.lastname, dbexcel.birthday, dbexcel.roll, dbexcel.caste, dbexcel.gender, dbexcel.email, dbexcel.phone, dbexcel.year, dbexcel.branch, dbexcel.position, dbexcel.coursename];


            connection.query(sql,[[records]], function (err, row, fields) {
                if (err) {
                    console.log("error is there")
                    throw err;
                } else {
                    console.log("SQL query done")
                }

            })

            res.redirect("/loginsuccess");

        } else {
            res.send("<h1>You have already filled this form!</h1>");
        }
    })

});

app.get("/loginsuccess", function (req, res) {
    let conn;
    var data1;

    connection.query("SELECT * FROM registration", function (err, students, fields) {

        const jsonStudents = JSON.parse(JSON.stringify(students));
        //   console.log(jsonStudents);

        console.log("I am in Loginsuccess function")
        let workbook = new excel.Workbook();
        let worksheet = workbook.addWorksheet('Customers'); //creating worksheet

        worksheet.columns = [
            { header: 'Name', key: 'Name', width: 10 },
            { header: 'Lastname', key: 'Lastname', width: 10 },
            { header: 'DOB', key: 'DOB', width: 10 },
            { header: 'Rollno', key: 'Rollno', width: 10 },
            { header: 'Caste', key: 'Caste', width: 10 },
            { header: 'Gender', key: 'Gender', width: 10 },
            { header: 'Email', key: 'Email', width: 10 },
            { header: 'Phone', key: 'Phone', width: 10 },
            { header: 'Year', key: 'Year', width: 10 },
            { header: "Branch", key: 'Branch', width: 20 },
            { header: 'Student/Faculty', key: 'Student/Faculty', width: 10 },
            { header: 'Coursename', key: 'Coursename', width: 10 },
        ];

        worksheet.addRows(jsonStudents);
        workbook.xlsx.writeFile("customer.xlsx")
            .then(function () {
                console.log("file saved!");
            });
    });
    res.redirect("/done.html");

});


app.listen(app.get('port'), function () {
    console.log("Server is running on port 8081");
});