// 后端的API接口
import { post } from './http'

const api_request = {
    File: {
        upload(data: any) {
            return post('/uploadpdf/', data)
        },
    }
}

export default api_request