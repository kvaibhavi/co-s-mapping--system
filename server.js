var express=require("express")
var flatten = require('flat').flatten;
var bodypar=require("body-parser")
var urlen=bodypar.urlencoded({extended:true})
var app=express()
var mongoose=require("mongoose")
var router=express.Router();
var jsdom=require("jsdom").JSDOM;
const { JSDOM } = jsdom;
var fs=require("fs");
var got=require("got")
var mongoxlsx=require("mongo-xlsx")
var excel=require("exceljs")
var xl=require('excel4node')
const vgmUrl='http://localhost:3007/faculty.html'
mongoose.connect('mongodb://127.0.0.1:27017/4thsempro',function(err,db)
{
    if(err)
        console.log("not connected");
    console.log("connected");


 login_data=mongoose.model('login',mongoose.Schema({
    user:{
        type:String,
        require:true
    },
    password:{
        type:String,
        require:true
    }
}))
})

course_data=mongoose.model('student_course_select',mongoose.Schema({
    usn:String,
    name:String,   
    
    core1:String,
    core2:String,
    elective1:String,
    elective2:String,
    elective3:String,
    miniproject:String

}))

co_excel=mongoose.model('co_mapping',mongoose.Schema({
    subject:String,
    test:String,
    Qno:String,
    maxMarks:Number,
    CO:String
}))

app.use(bodypar.urlencoded({extended:true}));
app.use(bodypar.json());
app.use("/mystyles",express.static('mystyles'));
var courses;

app.get("/",function(req,res)
{
    res.sendFile(__dirname+"/login.html");
})
app.get("/#back",function(req,res){
    res.sendFile(__dirname+"/admin.html");
})
router.post("/login_post",urlen,function(req,res,next){

    uname=req.body['un'];
    pass=req.body['pwd']
    login_data.findOne({user:uname,password:pass},function(err,data){
        if(err)
            res.send("error")
        if(data!=null)
        {
            if(uname=="admin")
                res.sendFile(__dirname+'/admin.html')
            else
            {       
             res.sendFile(__dirname+'/faculty.html')  
            }
        }
        else
        {
            
        res.sendFile(__dirname+'/login.html')  
        }      
    })    
})


app.post("/admin_batch_post",urlen,function(req,res,next)
{
    batch=req.body.batch;
    if(batch=="2018")
        res.sendFile(__dirname+"/admin_course.html")
})



router.post("/admin_course_post",urlen,function(req,res){
    course_records=new course_data({
        usn:req.body.usn,
        name:req.body.name1,
        core1:req.body.core1,
        core2:req.body.core2,
        elective1:req.body.elective1,
        elective2:req.body.elective2,
        elective3:req.body.elective3,
        miniproject:req.body.minipro
    });
    
    course_data.create(course_records,function(err,data)
    {
        if(err)
            res.send("error");
            res.sendFile(__dirname+"/admin_course.html")
    })
   /* course_records.save(function(err){
        if(err)
            res.send("error");
            res.sendFile(__dirname+"/admin_course.html")
    })*/

})

router.post("/faculty_question_post",urlen,function(req,res){
    co_records=new co_excel({
        subject:req.body.course,
        test:req.body.test,
        Qno:req.body.qno,
        maxMarks:req.body.maxMarks,
        CO:req.body.cos    
    });
    
    co_excel.create(co_records,function(err,data)
    {
        if(err)
            res.send("error");
            res.sendFile(__dirname+"/faculty.html")
    })
})


app.post('/generate_post',urlen,function(req,res){
    res.sendFile(__dirname+"/generate.html")
})


router.post('/generate_question_post',urlen,function(req,res){
    courses_id=req.body.course;
    test=req.body.test;
    fileName=req.body.fname;
    global.titleColumn=3;
    var wb=new xl.Workbook();
    var ws=wb.addWorksheet("sheet 2");
    var myStyle = wb.createStyle({
    alignment: {
        horizontal: 'center',
        wrapTesxt:true
      },
      font:{
          color:'#00FFFF',
          size:15,
          bold:true
      }
    });

    /*var myStyle1 = wb.createStyle({
        alignment: {
            horizontal: 'center',
            wrapTesxt:true
          },
          font:{
              color:'#00FFFF',
              size:15,
              bold:true
          },
          sheetProtection:{
            sort:true
          }
        });*/

    var secstyle=wb.createStyle({
        alignment:{
            horizontal: 'center',
            wrapTesxt:true
        },
        font:{
           // bold:true
        }
    })

    let rowIndex = 5;
    let titleRow=rowIndex-1;
    
    if(courses_id==="18mca41")
    {
        course_data.find({core1:courses_id}).stream()
        .on('data',function(doc)
        {
            var columnIndex1;
            co_excel.find({subject:courses_id}).stream()
            .on('data',function(doc2){
                
                if(doc2.test==="Test1")      
                    columnIndex1=3;
                else if (doc2.test==="Test2")
                    columnIndex1=8;
                else
                    console.log("wrong column")


            co_excel.find({subject:courses_id}).stream()
            .on('data',function(doc1){
                
            
                let rowIndex1 = 2;
                //ws.cell(1,3,2,6,true).string(doc1.test).style(myStyle)
                ws.cell(1,3,1,10,true).string(test).style(myStyle)
                ws.cell(rowIndex1++,columnIndex1).string(doc1.Qno).style(secstyle)
                ws.cell(rowIndex1++,columnIndex1).number(doc1.maxMarks).style(secstyle)
                ws.cell(rowIndex1++,columnIndex1).string(doc1.CO).style(secstyle)
                columnIndex1++;
                
                wb.write(fileName)
            
            })
        })
            //ws.cell(4,proxy1).string("Total").style(secstyle)
     
            let columnIndex = 1;
            ws.cell(titleRow,12).string("Total").style(secstyle)
           // ws.cell(1,3,2,6,true).string("Student Details").style(myStyle)
           ws.cell(titleRow,1).string("USN").style(secstyle)
           ws.cell(titleRow,2).string("NAME").style(secstyle)
            ws.column(columnIndex).setWidth(17);
        
            ws.cell(rowIndex,columnIndex++)
                        .string(doc.usn)
            ws.column(columnIndex).setWidth(20);
            ws.cell(rowIndex,columnIndex++)
                        .string(doc.name)
            rowIndex++;        
            wb.write(fileName)
            res.end()
        })      
    }
    else if(courses_id==="18mca42")
        {
            course_data.find({core2:courses_id}).stream()
            .on('data',function(doc)
            {
        
                let columnIndex1=3;
                co_excel.find({subject:courses_id,test:test}).stream()
                .on('data',function(doc1){
                    
                    let rowIndex1 = 2;
                    //ws.cell(1,3,2,6,true).string(doc1.test).style(myStyle)
                    ws.cell(1,3,1,10,true).string(test).style(myStyle)
                    ws.cell(rowIndex1++,columnIndex1).string(doc1.Qno).style(secstyle)
                    ws.cell(rowIndex1++,columnIndex1).number(doc1.maxMarks).style(secstyle)
                    ws.cell(rowIndex1++,columnIndex1).string(doc1.CO).style(secstyle)
                    columnIndex1++;
        
                    wb.write(fileName)
                })
                let columnIndex = 1;
               // ws.cell(1,3,2,6,true).string("Student Details").style(myStyle)
               ws.cell(titleRow,1).string("USN").style(secstyle)
               ws.cell(titleRow,2).string("NAME").style(secstyle)
                ws.column(columnIndex).setWidth(17);
            
                ws.cell(rowIndex,columnIndex++)
                            .string(doc.usn)
                ws.column(columnIndex).setWidth(20);
                ws.cell(rowIndex,columnIndex++)
                            .string(doc.name)
                rowIndex++;        
                wb.write(fileName)
                res.end()
            })      
        }
    else if(courses_id==="18mca431" || courses_id==="18mca432" || courses_id==="18mca433")
        {
            course_data.find({elective1:courses_id}).stream()
            .on('data',function(doc)
            {
             let rowIndex1 = 2;
               
                let columnIndex1=3;
                co_excel.find({subject:courses_id,test:test}).stream()
                .on('data',function(doc1){
                    
                    //ws.cell(1,3,2,6,true).string(doc1.test).style(myStyle)
                    ws.cell(1,3,1,10,true).string(test).style(myStyle)
                    ws.cell(rowIndex1++,columnIndex1).string(doc1.Qno).style(secstyle)
                    ws.cell(rowIndex1++,columnIndex1).number(doc1.maxMarks).style(secstyle)
                    ws.cell(rowIndex1++,columnIndex1).string(doc1.CO).style(secstyle)
                    columnIndex1++;
        
                    wb.write(fileName)
                })
                let columnIndex = 1;
               // ws.cell(1,3,2,6,true).string("Student Details").style(myStyle)
               ws.cell(titleRow,1).string("USN").style(secstyle)
               ws.cell(titleRow,2).string("NAME").style(secstyle)
                ws.column(columnIndex).setWidth(17);
            
                ws.cell(rowIndex,columnIndex++)
                            .string(doc.usn)
                ws.column(columnIndex).setWidth(20);
                ws.cell(rowIndex,columnIndex++)
                            .string(doc.name)
                rowIndex++;        
                wb.write(fileName)
                res.end()
            })      
        }

        else if(courses_id==="18mca441" || courses_id==="18mca442" || courses_id==="18mca443")
        {
            course_data.find({elective2:courses_id}).stream()
            .on('data',function(doc)
            {
        
                let columnIndex1=3;
                co_excel.find({subject:courses_id,test:test}).stream()
                .on('data',function(doc1){
                    
                    let rowIndex1 = 2;
                    //ws.cell(1,3,2,6,true).string(doc1.test).style(myStyle)
                    ws.cell(1,3,1,10,true).string(test).style(myStyle)
                    ws.cell(rowIndex1++,columnIndex1).string(doc1.Qno).style(secstyle)
                    ws.cell(rowIndex1++,columnIndex1).number(doc1.maxMarks).style(secstyle)
                    ws.cell(rowIndex1++,columnIndex1).string(doc1.CO).style(secstyle)
                    columnIndex1++;
        
                    wb.write(fileName)
                })
                let columnIndex = 1;
               // ws.cell(1,3,2,6,true).string("Student Details").style(myStyle)
               ws.cell(titleRow,1).string("USN").style(secstyle)
               ws.cell(titleRow,2).string("NAME").style(secstyle)
                ws.column(columnIndex).setWidth(17);
            
                ws.cell(rowIndex,columnIndex++)
                            .string(doc.usn)
                ws.column(columnIndex).setWidth(20);
                ws.cell(rowIndex,columnIndex++)
                            .string(doc.name)
                rowIndex++;        
                wb.write(fileName)
                res.end()
            })      
        }

        else if(courses_id==="18mca451" || courses_id==="18mca452" || courses_id==="18mca453")
        {            
            course_data.find({elective3:courses_id}).stream()
            .on('data',function(doc)
            {
        
                let columnIndex1=3;
                co_excel.find({subject:courses_id,test:test}).stream()
                .on('data',function(doc1){
                    
                    let rowIndex1 = 2;
                    //ws.cell(1,3,2,6,true).string(doc1.test).style(myStyle)
                   
                    ws.cell(rowIndex1++,columnIndex1).string(doc1.Qno).style(secstyle)
                    ws.cell(rowIndex1++,columnIndex1).number(doc1.maxMarks).style(secstyle)
                    ws.cell(rowIndex1++,columnIndex1).string(doc1.CO).style(secstyle)
                    columnIndex1++;
                    
                    wb.write(fileName)
                })
                var titleColumn=0;
                co_excel.count({'subject':courses_id,'test':test},function(err,count)
                {
                    ws.cell(1,3,1,(count+3),true).string(test).style(myStyle)
                    ws.cell(titleRow,(count+3)).string("Total").style(secstyle)
                    
                })

                
                let columnIndex = 1;
               // ws.cell(1,3,2,6,true).string("Student Details").style(myStyle)
               
               ws.cell(titleRow,1).string("USN").style(secstyle)
               ws.cell(titleRow,2).string("NAME").style(secstyle)
                ws.column(columnIndex).setWidth(17);
            
                ws.cell(rowIndex,columnIndex++)
                            .string(doc.usn)
                ws.column(columnIndex).setWidth(20);
                ws.cell(rowIndex,columnIndex++)
                            .string(doc.name)
                rowIndex++;        
                wb.write(fileName)
                res.end()
            })      
        }


    else
        console.log("wrong")


})  

app.use("/",router)
app.listen(3005)
