import Ratting from "../models/rattingModels.js";

class RattingControllers {

    async create (req, res) {
        try {
            const {value, label} = req.body;
            if (!value || !label) {
                return res.status(403).json('Значения полей рейтинга, не должны быть пустыми')
            }
            const cat = await Ratting.create({value, label});
            return res.json({cat}).status(200);
        }catch (e) {
            res.status(500).json("Something went wrong")
        }
    }

    async getAll (req, res) {
        try {
            const cat = await Ratting.find()
            res.json(cat)
        }catch (e) {
            res.status(501).json('Не выполнено')
        };
    }

    async getOne (req, res) {
        try {
            const cat = await Ratting.findById({_id: req.params.id});
            if (!cat) {
                return res.status(404).json('рейтинг не найдена')
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
                return res.status(403).json('Значения полей рейтинга, не должны быть пустыми')
            }
            await Ratting.findById({_id});
            await Ratting.updateOne({value, label})
            return res.status(200).json(({value, label}))
        }catch (e) {
            res.status(501).json('Не выполнено')
        }
    }

    async delete (req, res) {
        try{
            await Ratting.findByIdAndDelete({_id: req.params.id});
            return res.json(`удален`).status(200)
        }catch (e){
            res.status(500).json(e)
        }
    }
}
export default new RattingControllers()