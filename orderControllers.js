import Order from "../models/orderModels.js";

class OrderControllers {

    // async create (req, res) {
    //
    // }
    //
    // async getAll (req, res) {
    //
    // }

    async getMine (req, res) {
        try {
            const order = await Order.findById({userId: req.params.id});
            if (!order) {
                return res.status(404).json('Пользователь не найден')
            };
            return res.json(order).status(200);
        }catch (e) {
            res.status(501).json('Не выполнено')
        }
    }

    async getOne (req, res) {
        try {
            const order = await Order.findById({_id: req.params.id});
            if (!order) {
                return res.status(404).json('Заказ не найден')
            };
            return res.json(order).status(200);
        }catch (e) {
            res.status(501).json('Не выполнено')
        }
    }

    // async update (req, res) {
    //
    // }

    async delete (req, res) {
        try{
            await Order.findByIdAndDelete({_id: req.params.id});
            return res.json(`удален`).status(200)
        }catch (e){
            res.status(500).json(e)
        }
    }
}

export default new OrderControllers()