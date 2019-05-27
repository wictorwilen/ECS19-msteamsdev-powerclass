import { RequestFactory } from "./RequestFactory";
export class ClientRequestFactory<T> extends RequestFactory<T> {

    public get(uri: string, token: string): Promise<T | T[]> {
        return new Promise<T>((resolve, reject) => {
            fetch(new Request(uri, {
                headers: new Headers([
                    ["Authorization", "Bearer " + token],
                    ["Accept", "application/json"],
                    ["Content-Type", "application/json; charset=utf-8"],
                    ["OData-MaxVersion", "4.0"],
                    ["OData-Version", "4.0"],
                    ["Prefer", 'odata.include-annotations="*"']
                ])
            })).then(response => {
                response.json().then((json: any) => {
                    resolve(json.value);
                });
            });
        });
    }

    public post(uri: string, token: string, data: any): Promise<T | T[]> {
        return new Promise<T>((resolve, reject) => {
            fetch(new Request(uri, {
                method: "POST",
                headers: new Headers([
                    ["Authorization", "Bearer " + token],
                    ["Accept", "application/json"],
                    ["Content-Type", "application/json; charset=utf-8"],
                    ["OData-MaxVersion", "4.0"],
                    ["OData-Version", "4.0"],
                    ["Prefer", 'odata.include-annotations="*"']
                ]),
                body: JSON.stringify(data)
            })).then(response => {
                response.json().then((json: any) => {
                    resolve(json.value);
                });
            });
        });
    }
    public postHeader(uri: string, token: string, data: any, header: string): Promise<string> {
        throw new Error("Not implemented");
    }
}
