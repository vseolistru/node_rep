import User from "../models/userModels.js";
import bcrypt from 'bcryptjs';
import {v4} from 'uuid'
import mailService from "../service/mailService.js";

const strongPassword = new RegExp("^(?=.*[a-z])(?=.*[A-Z])(?=.*[0-9])(?=.*[!@#$%^&*])(?=.{8,})");
const emailValidator = new RegExp("^([A-Za-z0-9_\\-\\.])+\\@([A-Za-z0-9_\\-\\.])+\\.([A-Za-z]{2,4})")

class UserControllers {
    async registration (req, res) {
        try {
            const {username, email, password} = req.body
            if (!email || !password || !username) {
                return res.status(403).json('Значения полей username, email и password не должны быть пустыми')
            }
            if (!emailValidator.test(email)) {
                return res.status(403).json('Email должен содержать - @ и домен')
            }
            if (password.length < 8) {
                return res.status(403).json('Пароль должен содержать 8 символов')
            }
            if (!strongPassword.test(password)) {
                return res.status(403).json('Пароль должен содержать 1 спецсимвол, 1 символ верхнего регистра, 1 число')
            }
            const candidate = await User.findOne({email})
            if (candidate) {
                return res.status(403).json('Пользователь с таким email уже существует')
            }
            const hashPassword = await bcrypt.hashSync(password)
            const activationLink = v4()
            await mailService.sendActivationMail(email, `${process.env.API_URL}/api/users/activate/${activationLink}`)
            const user = await User.create({email, username, activationLink,password: hashPassword})
            return res.json({user})
        }
        catch (e) {
            res.status(500).json("Something went wrong")
        }
    }

    async login (req, res) {
        try {
            const {email, password} = req.body
            const user = await User.findOne({email})
            if (!user){
                return res.status(401).json(`user ${email} does not exist`)
            }
            let comparePassword = bcrypt.compareSync(password, user.password)
            if (!comparePassword){
                return res.status(401).json(`user pass is a wrong`)
            }
            return res.status(200).json({_id:user._id, username: user.username, email: user.email})
        } catch (e) {
            res.status(501).json('Не выполнено')
        }
    }

    // async auth (req, res) {
    //
    // }

    async activate (req, res, next,){
        try {
            const activationLink = req.params.link
            const  userLink = await User.findOne({activationLink})
            if(!userLink){
                res.status(500).json('Некорректная ссылка')
            }
            userLink.isActivated = true
            await userLink.save()
            return res.redirect('http://localhost:5000/')
        }
        catch (e){
            res.status(500).json('is not saved')
        }
    }
    async getAll (req, res) {
        try {
            const users = await User.find()
            res.json(users)
        }catch (e) {
            res.status(501).json('Не выполнено')
        }
    }

    async getOne (req, res) {
        try {
            const user = await User.findById({_id: req.params.id})
            if (!user) {
                return res.status(404).json('Пользователь не найден')
            }
            return res.json(user)
        }catch (e) {
            res.status(501).json('Не выполнено')
        }
    }

    async update (req, res) {
        try {
            const {_id, email, password} = req.body
            const user = await User.findById({_id})
            if (!email || !password) {
                return res.status(403).json('Значения полей email и password не должны быть пустыми')
            }
            if (password.length < 8) {
                return res.status(403).json('Пароль должен содержать 8 символов')
            }
            if (!strongPassword.test(password)) {
                return res.status(403).json('Пароль должен содержать 1 спецсимвол, 1 символ верхнего регистра, 1 число')
            }
            const hashPassword = await bcrypt.hashSync(password)
            const updateUser = await User.updateOne({email,password: hashPassword})
            return res.json({_id, email, password: hashPassword}).status(200)
        } catch (e) {
            res.status(501).json('Не выполнено')
        }
    }

}

export default new UserControllers()