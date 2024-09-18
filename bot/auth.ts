import { ConfidentialClientApplication } from "@azure/msal-node";
import jwt from "jsonwebtoken";
import config from "./config";
import jwkToPem from "jwk-to-pem";

// MSAL configuration
const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        clientSecret: config.clientSecret,
    }
};

const cca = new ConfidentialClientApplication(msalConfig);

export const getAccessTokenOnBehalfOf = async (requestToken: string, targetScope: string): Promise<string> => {
    const oboRequest = {
        oboAssertion: requestToken,
        scopes: [targetScope]
    };

    try {
        const response = await cca.acquireTokenOnBehalfOf(oboRequest);
        return response.accessToken;
    } catch (error) {
        console.error('Error acquiring token on behalf of:', error);
        throw new Error('Failed to acquire token on behalf of');
    }
}

// Middleware to verify the JWT token
export const authenticate = async (req, res) => {
    // get the authorization header
    const authHeader = req.headers["authorization"];
    
    // ensure it exists and contains a bearer token
    if (!authHeader || !authHeader.startsWith("Bearer ")) {
        res.send(401, { message: "Authorization header missing or malformed" });
        return;
    }

    // split out the token
    const token = authHeader.split(" ")[1];

    // get public keys to verify token
    const response = await fetch(`https://login.microsoftonline.com/${config.tenantId}/discovery/v2.0/keys`);
    const { keys } = await response.json();
    const decodedHeader:any = jwt.decode(token, { complete: true });

    // find the key used to sign the token
    const key = keys.find(k => k.kid === decodedHeader.header.kid);

    // convert the key to pem format
    const pemKey = jwkToPem(key);

    // Verify the JWT token
    jwt.verify(token, pemKey, { algorithms: ["RS256"] }, (err, decoded) => {
        if (err) {
            res.send(403, { message: "Invalid or expired token" });
            return;
        }

        // Token is valid
        req.user = decoded;
        return;
    });
};

// Function to acquire an access token for API-to-API communication (if needed)
export const acquireTokenForApi = async (scopes: string[]) => {
    try {
        const result = await cca.acquireTokenByClientCredential({
            scopes: scopes
        });
        return result.accessToken;
    } catch (error) {
        console.error("Error acquiring token:", error);
        return null;
    }
};
