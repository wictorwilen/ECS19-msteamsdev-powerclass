import DynamicsEntity from "./DynamicsEntity";
import { RequestFactory } from "./RequestFactory";
import { Account } from "./DynamicsDefinitions";

export default class Accounts<S extends RequestFactory<Account>> extends DynamicsEntity<Account, S> {
    public entityName = "accounts";
    public entityIdName = "accountid";

    constructor(token: string, factory: S) {
        super(token, factory);
    }
}
