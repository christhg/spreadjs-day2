import axios from "axios"
import { Loading, Notification } from 'element-ui';

const instance = axios.create({
    baseURL: "http://localhost:8088",
    timeout: 10000,
    headers: {'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Headers': 'Origin, X-Requested-With, Content-Type, Accept',
            }
});
class HttpUtils {
    public static get(url: string, config: any = {}) {
        return new Promise((resolve, reject)=>{
            let loadingInstance = Loading.service({ fullscreen: true });
            instance.get(url, config)
            .then(response=>{ resolve(response.data)})
            .catch((err)=>{
                Notification({title: '错误',message: err,type:'error'});
                resolve(undefined);
            })
            .finally(()=>{ loadingInstance.close(); });
        });
    }

    public static post(url: string, params: any = {}, config:any = {}) {

        return new Promise((resolve, reject)=>{
            let loadingInstance = Loading.service({ fullscreen: true });
            instance.post(url, params, config)
            .then(response=>{ resolve(response.data)})
            .catch((err)=>{
                Notification({title: '错误',message: err,type:'error'});
                resolve(undefined);
            })
            .finally(()=>{ loadingInstance.close(); });
        });
    }
}

export default HttpUtils;