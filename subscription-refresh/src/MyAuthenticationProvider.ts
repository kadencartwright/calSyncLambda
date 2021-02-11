import { AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import axios from 'axios'
import * as qs from 'qs'
export class MyAuthenticationProvider implements AuthenticationProvider {
	public token:Token = new Token('invalidToken',0);
	/**
	 * This method will get called before every request to the msgraph server
	 * This should return a Promise that resolves to an accessToken (in case of success) or rejects with error (in case of failure)
	 * Basically this method will contain the implementation for getting and refreshing accessTokens
	 */ 
	public async getAccessToken(): Promise<string> {
		if (this.token.isValid()){
			return this.token.bearer;
		}else{
			try{
				this.token = await getToken()
				return this.token.bearer
			}catch(e){
				console.log(e)
			}
		}
	}
};


class Token{
	expires_in_milliseconds:number;
	created_at:Date;

	bearer:string;
	constructor(bearer:string,expiresInSeconds:number){
		this.expires_in_milliseconds= (expiresInSeconds-30)*1000;//if token expires in next 30 seconds we will go ahead and get another one
		this.bearer= bearer;
		this.created_at = new Date();//record the creation time of the current token
	}

	isValid: ()=>boolean = function(){
		if(new Date().getTime() <(this.created_at.getTime()+this.expires_in_milliseconds)){
			return true;
		}else{
			return false
		}
	}
}
let getToken:()=>Promise<Token> = async function(){
    let url =`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`
    let params = {
        client_secret: process.env.CLIENT_SECRET,
        grant_type: 'client_credentials',
        tenant: process.env.TENANT_ID,
        client_id: process.env.CLIENT_ID,
        scope: 'https://graph.microsoft.com/.default'
	  }
	let response = await axios({
		method: 'post',
		url: url,
		data: qs.stringify(params),
		headers: {
		  'content-type': 'application/x-www-form-urlencoded;charset=utf-8'
		}
	})
   
    if(!response.data.access_token){
        throw new Error('failed to get access token')
    }else{
        return new Token(response.data.access_token,response.data.expires_in);
    }
}
