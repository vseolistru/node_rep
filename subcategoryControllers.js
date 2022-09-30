import Subcategory from "../models/subcategoryModels.js";

class SubcategoryControllers {
    async create (req, res) {
        try {
            const {checked, label} = req.body;
            if (!checked || !label) {
                return res.status(403).json('Значения полей категории, не должны быть пустыми')
            }
            const subCat = await Subcategory.create({checked, label});
            return res.json({subCat}).status(200);
        }catch (e) {
            res.status(500).json("Something went wrong")
        }
    }

    async getAll (req, res) {
        try {
            const subCat = await Subcategory.find()
            res.json(subCat)
        }catch (e) {
            res.status(501).json('Не выполнено')
        };
    }

    async getOne (req, res) {
        try {
            const subCat = await Subcategory.findById({_id: req.params.id});
            if (!subCat) {
                return res.status(404).json('Категория не найдена')
            };
            return res.json(subCat);
        }catch (e) {
            res.status(501).json('Не выполнено')
        }
    }

    async update (req, res) {
        try {
            const {_id, checked, label} = req.body;
            if (!checked || !label){
                return res.status(403).json('Значения полей категории, не должны быть пустыми')
            }
            await Subcategory.findById({_id});
            await Subcategory.updateOne({checked, label})
            return res.status(200).json(({checked, label}))
        }catch (e) {
            res.status(501).json('Не выполнено')
        }
    }

    async delete (req, res) {
        try{
            await Subcategory.findByIdAndDelete({_id: req.params.id});
            return res.json(`удален`).status(200)
        }catch (e){
            res.status(500).json(e)
        }
    }

}

export default new SubcategoryControllers()