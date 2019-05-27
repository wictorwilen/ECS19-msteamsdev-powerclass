import { RequestFactory } from "./RequestFactory";

// tslint:disable: no-angle-bracket-type-assertion

export default abstract class DynamicsEntity<T, S extends RequestFactory<T>> {

    public abstract readonly entityName: string;
    public abstract readonly entityIdName: string;


    constructor(private token: string, private factory: S) {
    }

    public async getById(id: string): Promise<T> {
        return <Promise<T>> this.factory.get(
            decodeURI(`https://${process.env.TENANT_NAME}.crm4.dynamics.com/api/data/v9.0/${this.entityName}(${id})`),
            this.token);
    }

    public async getAll(count: number = 10): Promise<T[]> {
        return <Promise<T[]>> this.factory.get(
            decodeURI(`https://${process.env.TENANT_NAME}.crm4.dynamics.com/api/data/v9.0/${this.entityName}?$top=${count}`),
            this.token);

    }
    public async getByFilter(filter: string, expand?: string, count?: number): Promise<T[]> {
        let uri = `https://${process.env.TENANT_NAME}.crm4.dynamics.com/api/data/v9.0/${this.entityName}?$filter=${filter}`;
        if (count) {
            uri = uri + `&$top=${count}`;
        }
        return <Promise<T[]>> this.factory.get(
            decodeURI(uri),
            this.token);
    }
    public async getSubEntity<E>(filter: string, linkedEntity: string): Promise<E[]> {
        const uri = `https://${process.env.TENANT_NAME}.crm4.dynamics.com/api/data/v9.0/${this.entityName}?$expand=${linkedEntity}&$filter=${filter}`;

        const x = await <Promise<T[]>> this.factory.get(
            decodeURI(uri),
            this.token);
        if (x.length === 0) {
            return [];
        }
        return x[0][linkedEntity];
    }
    public async getGenericById(id: string, expand: string): Promise<any> {
        const uri = `https://${process.env.TENANT_NAME}.crm4.dynamics.com/api/data/v9.0/${this.entityName}(${id})/?$expand=${expand}`;
        return <Promise<any>> this.factory.get(
            decodeURI(uri),
            this.token);
    }

    public async resolve(body: any): Promise<any> {
        const uri = `https://${process.env.TENANT_NAME}.crm4.dynamics.com/api/data/v9.0/ResolveIncident?tag=abortbpf`;
        return <Promise<any>> this.factory.post(
            decodeURI(uri),
            this.token,
            body);
    }

    public async add(data: Partial<T>): Promise<T> {
        const uri = `https://${process.env.TENANT_NAME}.crm4.dynamics.com/api/data/v9.0/${this.entityName}`;
        return new Promise<T>((resolve, reject) => {
            this.factory.postHeader(
                decodeURI(uri),
                this.token,
                data,
                "odata-entityid").then(entityPath => {
                    this.factory.get(entityPath, this.token).then((entity: T) => {
                        resolve(entity);
                    }).catch(err => {
                        reject(err);
                    });
                }).catch(err => {
                    reject(err);
                });
        });
    }
}
