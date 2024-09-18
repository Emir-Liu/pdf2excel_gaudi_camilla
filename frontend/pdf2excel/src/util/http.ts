
import axios from "axios"

// 后端地址
export const DOMAIN = '127.0.0.1:8000'
export const baseUrl = 'http://' + DOMAIN

axios.defaults.baseURL = baseUrl

// 对axio进行配置和封装
// GET请求
export function get(url: string, params: object) {
    // console.log('param:',params)
    return axios.get(url, { params })
        .then(function (res) {
            return res
        })
        .catch(function () {
            // 异常处理
        })
}

// POST请求
// 需要注意的是，POST请求发送的使用DataFrame，因为支持文件发送
export function post(url: string, params: object) {
    return axios.post(url, params)
        .then(function (res) { return res })
        .catch(function () {
            // 异常处理
        })
}
