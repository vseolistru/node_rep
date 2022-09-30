import bcrypt from 'bcryptjs'

const data = {
    users: [
        {
            username: "Admin",
            email: "somemail@test.com",
            password: bcrypt.hashSync("somepass"),
            isAdmin: true
        }
    ],
    blogs:[
        {
            title: "Blog one",
            description : "Lorem ipsum dolor sit amet consectetur",
        },
        {
            title: "Blog two",
            description : "Lorem ipsum dolor sit amet consectetur",
        },
        {
            title: "Blog three",
            description : "Lorem ipsum dolor sit amet consectetur",
        },
        {
            title: "Blog four",
            description : "Lorem ipsum dolor sit amet consectetur",
        },
        {
            title: "Blog five",
            description : "Lorem ipsum dolor sit amet consectetur",
        }
    ],
    products: [
        {
            title: "Product one",
            category: "Women",
            subcategory: "Dresses",
            description : "Some description",
            price : "200.00",
            star: "3*",
            sizes: [
                {title:"M"},
                {title:"S"},
                {title:"L"},
                {title:"XL"},
            ],
            colors :[
                {title:"Blue"},
                {title:"Red"},
                {title:"White"},
            ],
            image: "./assets/images/products/1.0.webp",
            imageOne: "./assets/images/products/1.1.webp"
        },
        {
            title: "Product two",
            category: "Men",
            subcategory: "Jeans",
            description : "Some jeans description",
            price : "90.00",
            star: "4*",
            sizes: [
                {title:"M"},
                {title:"S"},
                {title:"L"},
                {title:"XL"},
            ],
            colors :[
                {title:"Blue"},
                {title:"Red"},
                {title:"White"},
            ],
            image: "./assets/images/products/2.0.webp",
            imageOne: "./assets/images/products/2.1.webp"
        },
        {
            title: "Product three",
            category: "Men",
            subcategory: "Jeans",
            description : "Some jeans description",
            price : "94.00",
            star: "4*",
            sizes: [
                {title:"M"},
                {title:"S"},
                {title:"L"},
                {title:"XL"},
            ],
            colors :[
                {title:"Blue"},
                {title:"Red"},
                {title:"White"},
            ],
            image: "./assets/images/products/3.0.webp",
            imageOne: "./assets/images/products/3.1.webp"
        },
        {
            title: "Product four",
            category: "Women",
            subcategory: "Dresses",
            description : "Some dress description",
            price : "180.00",
            star: "4*",
            sizes: [
                {title:"M"},
                {title:"S"},
                {title:"L"},
                {title:"XL"},
            ],
            colors :[
                {title:"Blue"},
                {title:"Red"},
                {title:"White"},
            ],
            image: "./assets/images/products/4.0.webp",
            imageOne: "./assets/images/products/4.1.webp"
        },
        {
            title: "Product five",
            category: "Men",
            subcategory: "T-shirt",
            description : "Some t-shirt description",
            price : "20.00",
            star: "4*",
            sizes: [
                {title:"M"},
                {title:"S"},
                {title:"L"},
                {title:"XL"},
            ],
            colors :[
                {title:"Blue"},
                {title:"Red"},
                {title:"White"},
            ],
            image: "./assets/images/products/5.0.webp",
            imageOne: "./assets/images/products/5.1.webp"
        },
    ],
    categories: [
        {value: "Men", label: "Men"},
        {value: "Women", label: "Women"},
        {value: "Kids", label: "Kids"}
    ],
    subcategory: [
        {checked: false, label: "Dresses"},
        {checked: false, label: "T-shirt"},
        {checked: false, label: "Jeans"},
        {checked: false, label: "Shorts"},
        {checked: false, label: "Skirts"}
    ],
    ratting:[
        {value: '1', label:"1*"},
        {value: '2', label:"2*"},
        {value: '3', label:"3*"},
        {value: '4', label:"4*"},
        {value: '5', label:"5*"}
    ],
}
export default data;

