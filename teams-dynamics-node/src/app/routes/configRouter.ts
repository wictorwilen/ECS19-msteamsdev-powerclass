import * as express from "express";
import JsonDB = require("node-json-db");
import passport = require("passport");
import * as debug from "debug";
// tslint:disable-next-line: no-var-requires
const BearerStrategy = require("passport-azure-ad").BearerStrategy;

const log = debug("msteams");

export const settingsRouter = (options: any): express.Router => {
    const router = express.Router();


    // set up auth for this endpoint
    const authOptions = {
        identityMetadata: "https://login.microsoftonline.com/common/.well-known/openid-configuration/",
        clientID: process.env.CLIENT_APP_ID,
        validateIssuer: false, // must be false as we do the common endpoint above
        loggingLevel: "warn", // info, warn, error
        passReqToCallback: false,
        audience: `https://${process.env.TENANT_NAME}.crm4.dynamics.com`, // cheating a bit by reusing the same aud (should really be the clientid)
        //loggingNoPII: false
    };

    const bearerStrategy = new BearerStrategy(authOptions,
        (token, done) => {
            // Send user info using the second argument
            done(null, {}, token);
        }
    );

    router.use(passport.initialize());
    passport.use(bearerStrategy);

    // Disable cache and allow CORS
    router.use((req, res, next) => {
        res.header("Cache-Control", "no-store, no-cache, must-revalidate, private")
        res.header("Access-Control-Allow-Origin", "*");
        res.header("Access-Control-Allow-Headers", "Authorization, Origin, X-Requested-With, Content-Type, Accept");
        next();
    });


    // get the settings
    router.get("/:tenant/:team/:channel",
        passport.authenticate("oauth-bearer", { session: false }),
        (req, res) => {
            log((req as any).user);
            const tenantId = req.params.tenant;
            const teamId = req.params.team;
            const channelId = req.params.channel;
            let setting: any;
            const settings = new JsonDB("settings", true, false);
            try {
                setting = settings.getData(`/${tenantId}/${teamId}/${channelId}`);
                res.json({ value: setting });
            } catch (err) {
                res.sendStatus(404);
            }
        });

    // Add a new setting and return the new settings
    router.post("/:tenant/:team/:channel",
        passport.authenticate("oauth-bearer", { session: false }),
        (req, res) => {
            const tenantId = req.params.tenant;
            const teamId = req.params.team;
            const channelId = req.params.channel;
            let setting: any;
            const settings = new JsonDB("settings", true, false);
            try {
                settings.push(`/${tenantId}/${teamId}/${channelId}`, req.body, false);
                setting = settings.getData(`/${tenantId}/${teamId}/${channelId}`);
                res.json({ value: setting });
            } catch (err) {
                res.sendStatus(500);
            }
        });

    return router;
}