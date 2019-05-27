import * as request from "request";
import { RequestFactory } from "./RequestFactory";
import * as debug from "debug";

const log = debug("msteams");
export class NodeRequestFactory<T> extends RequestFactory<T> {
    public get(uri: string, token: string): Promise<T | T[]> {
        return new Promise<T>((resolve, reject) => {
            try {
                request({
                    method: "GET",
                    uri,
                    headers: {
                        "Authorization": "Bearer " + token,
                        "Accept": "application/json",
                        "Content-Type": "application/json; charset=utf-8",
                        "OData-MaxVersion": "4.0",
                        "OData-Version": "4.0",
                        "Prefer": 'odata.include-annotations="*"'
                    },
                }, (error: any, response: any, body: any) => {
                    if (error) {
                        log(error);
                        reject(error);
                    } else if (response.statusCode === 200) {
                        const json = JSON.parse(body);
                        if (json.value) {
                            resolve(json.value);
                        } else {
                            resolve(json);
                        }
                    } else {
                        reject(response.statusCode);
                    }
                });
            } catch (err) {
                // tslint:disable-next-line: no-console
                console.log(err);
            }
        });
    }

    public post(uri: string, token: string, data: any): Promise<T | T[]> {
        return new Promise<T>((resolve, reject) => {
            try {
                request({
                    method: "POST",
                    uri,
                    headers: {
                        "Authorization": "Bearer " + token,
                        "Accept": "application/json",
                        "Content-Type": "application/json; charset=utf-8",
                        "OData-MaxVersion": "4.0",
                        "OData-Version": "4.0",
                        "Prefer": 'odata.include-annotations="*"'
                    },
                    body: JSON.stringify(data)
                }, (error: any, response: any, body: any) => {
                    if (error) {
                        log(error);
                        reject(error);
                    } else {
                        if (response.statusCode === 200 || response.statusCode === 204) {
                            resolve(undefined);
                        } else {
                            reject(response.statusCode);
                        }
                    }
                });
            } catch (err) {
                // tslint:disable-next-line: no-console
                log(err);
            }
        });
    }
    public postHeader(uri: string, token: string, data: any, header: string): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            try {
                request({
                    method: "POST",
                    uri,
                    headers: {
                        "Authorization": "Bearer " + token,
                        "Accept": "application/json",
                        "Content-Type": "application/json; charset=utf-8",
                        "OData-MaxVersion": "4.0",
                        "OData-Version": "4.0",
                        "Prefer": 'odata.include-annotations="*"'
                    },
                    body: JSON.stringify(data)
                }, (error: any, response: any, body: any) => {
                    if (error) {
                        log(error);
                        reject(error);
                    } else {
                        if (response.statusCode === 200 || response.statusCode === 204) {
                            resolve(response.headers[header]);
                        } else {
                            reject(response.statusCode);
                        }
                    }
                });
            } catch (err) {
                // tslint:disable-next-line: no-console
                log(err);
            }
        });
    }
}
