const express = require('express');
const cors=require('cors');
const router = require('./route/route');
require('dotenv').config()
const app = express();
app.use(cors())
const PORT=process.env.PORT || 4000
app.use(router)

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
