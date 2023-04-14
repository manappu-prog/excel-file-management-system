const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const mongoose = require('mongoose');
const port = 3000;

const app = express();

app.set('view engine', 'ejs');
app.set('views', 'views');

mongoose.connect('mongodb://127.0.0.1:27017/mydatabase', {
    useNewUrlParser: true,
    useUnifiedTopology: true
})
    .then(() => {
        console.log('Connected to MongoDB successfully!');
    })
    .catch(err => {
        console.error('Failed to connect to MongoDB:', err);
        process.exit(1);
    });


const UserSchema = new mongoose.Schema({
    rollnumber: String,
    name: String,
    phonenumber: String
});

const User = mongoose.model('User', UserSchema);

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'uploads/');
    },
    filename: function (req, file, cb) {
        cb(null, file.originalname);
    }
});

const upload = multer({ storage: storage });

app.get('/', (req, res) => {
    res.render('upload-form.ejs', { title: 'File Upload' });
})

app.get('/dashboard',(req,res) => {
    User.find({},function(err,data){
        if(err) res.redirect('back');
        res.render('dashboard.ejs',{data: data});
        
    })
})

app.post('/uploads', upload.single('excelFile'), function (req, res) {
    if (!req.file) {
        res.status(400).send('No file uploaded..');
    }

    const workbook = xlsx.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);


    User.insertMany(sheetData)
        .then(() => {
            console.log('Data stored in MongoDB successfully!');
            mongoose.connection.close();
        })
        .catch(err => {
            console.error('Failed to store data in MongoDB:', err);
            mongoose.connection.close();
        });

    res.status(200).send('File uploaded succsessfuly');
})

app.listen(port, (err) => {
    if (err) {
        console.log('error while running the server --> ' + err);
        return;
    }
    console.log('express app is up and running...');
})