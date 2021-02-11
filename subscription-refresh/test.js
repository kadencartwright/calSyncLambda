const express = require('express')

const app = express()
app.use(express.query())
app.use(express.urlencoded({extended:true}))

app.all('/*',(req,res)=>{
    res.type('text/plain')
    console.log(req.params)
    console.log(req.body)
    
    res.status(200).send(req.query.validationToken)
})



app.listen(7071,()=>{
    console.log('server up!')
})