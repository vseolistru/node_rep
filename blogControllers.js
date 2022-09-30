import Blog from "../models/blodModels.js";

class BlogControllers {

    async add (req, res) {
        try {
            const {title, description} = req.body;
            if (!title || !description) {
                return res.status(403).json('Значения полей статьи, не должны быть пустыми')
            }
            const article = await Blog.create({title, description, author:"Admin"});
            return res.json({article}).status(200);
        }catch (e) {
            res.status(500).json("Something went wrong")
        }
    }

    async getAll (req, res) {
        try {
            const blogs = await Blog.find()
            res.json(blogs)
        }catch (e) {
            res.status(501).json('Не выполнено')
        };
    }

    async getOne (req, res) {
        try {
            const blog = await Blog.findById({_id: req.params.id});
            if (!blog) {
                return res.status(404).json('статья не найдена')
            };
            return res.json(blog);
        }catch (e) {
            res.status(501).json('Не выполнено')
        }
    }
    async update (req, res) {
        try {
            const {_id, title, description} = req.body;
            if (!title || !description){
                return res.status(403).json('Значения полей статьи, не должны быть пустыми')
            }
            await Blog.findByIdAndUpdate({_id, title, description});
            return res.status(200).json(({title, description}))
        }catch (e) {
            res.status(501).json('Не выполнено')
        }

    }
    async delete (req, res) {
        try{
            await Blog.findByIdAndDelete({_id: req.params.id});
            return res.json(`удален`).status(200)
        }catch (e){
            res.status(500).json(e)
        }
    }
}

export default new BlogControllers();

