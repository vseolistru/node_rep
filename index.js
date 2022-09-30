import express from 'express'
import mongoose from 'mongoose'
import dotenv from 'dotenv'
import seedRoute from "./routes/seedRoute.js";
import userRoute from "./routes/userRoute.js";
import blogRoute from "./routes/blogRoutes.js";
import productRoute from "./routes/productRoute.js";
import subcategoryRoute from "./routes/subcategoryRoute.js";
import categoryRouter from "./routes/categoryRouter.js";
import rattingRoute from "./routes/rattingRoute.js";
import orderRoute from "./routes/orderRoute.js";



dotenv.config()
const PORT = process.env.PORT || 8000;
const app = express();
const DB_URL = process.env.DB_URL
app.use(express.json())
app.use(express.urlencoded({extended: true}))

app.use('/api/seed', seedRoute);
app.use('/api/users', userRoute);
app.use('/api/blogs', blogRoute);
app.use('/api/products', productRoute);
app.use('/api/cat', categoryRouter);
app.use('/api/subcat', subcategoryRoute);
app.use('/api/ratting', rattingRoute);
app.use('/api/orders', orderRoute);

app.get('/', (req, res)=>{
    res.status(200).json({message:'all fine'})
 });

async function startapp() {
    try{
        await mongoose.connect(DB_URL)
        app.listen(PORT, ()=> console.log(`success ${PORT}`));
    } catch (e){
        console.log(e)
    }
}

startapp()