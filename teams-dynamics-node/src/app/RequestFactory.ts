export abstract class RequestFactory<T> {
    public abstract get(uri: string, token: string): Promise<T | T[]>;
    public abstract post(uri: string, token: string, data: any): Promise<T | T[]>;
    public abstract postHeader(uri: string, token: string, data: any, header: string): Promise<string>;
}
