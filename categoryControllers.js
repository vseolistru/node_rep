import Category from "../models/categoryModels.js";

class CategoryControllers {

    async create (req, res) {
        try {
            const {value, label} = req.body;
            if (!value || !label) {
                return res.status(403).json('Значения полей Категории, не должны быть пустыми')
            }
            const cat = await Category.create({value, label});
            return res.json({cat}).status(200);
        }catch (e) {
            res.status(500).json("Something went wrong")
        }
    }

    async getAll (req, res) {
        try {
            const cat = await Category.find()
            res.json(cat)
        }catch (e) {
            res.status(501).json('Не выполнено')
        };
    }

    async getOne (req, res) {
        try {
            const cat = await Category.findById({_id: req.params.id});
            if (!cat) {
                return res.status(404).json('категория не найдена')
            };
            return res.json(cat);
        }catch (e) {
            res.status(501).json('Не выполнено')
        }
    }

    async update (req, res) {
        try {
            const {_id, value, label} = req.body;
            if (!value || !label){
                return res.status(403).json('Значения полей категории, не должны быть пустыми')
            }
            await Category.findById({_id});
            await Category.updateOne({value, label})
            return res.status(200).json(({value, label}))
        }catch (e) {
            res.status(501).json('Не выполнено')
        }
    }

    async delete (req, res) {
        try{
            await Category.findByIdAndDelete({_id: req.params.id});
            return res.json(`удален`).status(200)
        }catch (e){
            res.status(500).json(e)
        }
    }
}
export default new CategoryControllers()