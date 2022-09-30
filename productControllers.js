import Product from "../models/productModels.js";

class ProductController {

    async create (req, res) {
        try {
            const {title, category, subcategory, description, price, sizes, colors} = req.body;
            if (!title ||!category ||!subcategory ||!description ||!price ||!sizes ||!colors) {
                return res.status(403).json('Значения полей товаров, не должны быть пустыми')
            }
            const article = await Product.create({title, category, subcategory, description, price, sizes, colors, star:"2*"});
            return res.json({article}).status(200);
        }catch (e) {
            res.status(500).json("Something went wrong")
        }
    }

    async getAll (req, res) {
        try {
            const products = await Product.find()
            res.json(products).status(200)
        }catch (e) {
            res.status(501).json('Не выполнено')
        };

    }

    async getOne (req, res) {
        try {
            const product = await Product.findById({_id: req.params.id});
            if (!product) {
                return res.status(404).json('Пользователь не найден')
            };
            return res.json(product).status(200);
        }catch (e) {
            res.status(501).json('Не выполнено')
        }
    }

    async update (req, res) {
        try {
            const {_id,title, category, subcategory, description, price, sizes, colors, star} = req.body;
            if (!title ||!category ||!subcategory ||!description ||!price ||!sizes ||!colors ||!star) {
                return res.status(403).json('Значения полей товаров, не должны быть пустыми')
            }
            await Product.findById({_id});
            await Product.updateOne({title, category, subcategory, description, price, sizes, colors, star})
            return res.status(200).json(({_id,title, category, subcategory, description, price, sizes, colors, star}))
        }catch (e) {
            res.status(501).json('Не выполнено' )
        }
    }

    async delete (req, res) {
        try{
            await Product.findByIdAndDelete({_id: req.params.id});
            return res.json(`удален`).status(200)

        }catch (e){
            res.status(500).json(e)
        }
    }
}

export default new ProductController();