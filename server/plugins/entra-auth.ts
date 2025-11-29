import { FastifyInstance, FastifyRequest, FastifyReply } from "fastify";
import fastifyPlugin from "fastify-plugin";
import { createRemoteJWKSet, jwtVerify, JWTPayload } from "jose";

// Extend FastifyRequest to include user and token
declare module "fastify" {
  interface FastifyRequest {
    user?: JWTPayload;
    token?: string;
  }
  interface FastifyInstance {
    requireAuth: () => (
      request: FastifyRequest,
      reply: FastifyReply
    ) => Promise<void>;
    getToken: (request: FastifyRequest) => string | undefined;
  }
}

export default fastifyPlugin(async (fastify: FastifyInstance) => {
  const tenantId = process.env.ENTRA_TENANT_ID;
  const audience = process.env.ENTRA_AUDIENCE;
  const clientId = process.env.ENTRA_CLIENT_ID;

  if (!tenantId || !audience) {
    throw new Error(
      "ENTRA_TENANT_ID and ENTRA_AUDIENCE must be set in environment variables"
    );
  }

  // Accept both the api:// URI and the raw client ID as valid audiences
  const validAudiences = [audience];
  if (clientId && clientId !== audience) {
    validAudiences.push(clientId);
  }
  // Also add the client ID extracted from api:// URI if present
  if (audience.startsWith("api://")) {
    validAudiences.push(audience.replace("api://", ""));
  }

  // Accept both v1.0 and v2.0 token issuers
  const validIssuers = [
    `https://login.microsoftonline.com/${tenantId}/v2.0`,  // v2.0 issuer
    `https://sts.windows.net/${tenantId}/`,                // v1.0 issuer
  ];

  fastify.log.info(`Valid audiences: ${validAudiences.join(", ")}`);
  fastify.log.info(`Valid issuers: ${validIssuers.join(", ")}`);

  // Create JWKS client for token verification
  const JWKS = createRemoteJWKSet(
    new URL(
      `https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`
    )
  );

  // Decorate fastify with auth utilities
  fastify.decorate(
    "requireAuth",
    () => async (request: FastifyRequest, reply: FastifyReply) => {
      const authHeader = request.headers.authorization;

      if (!authHeader?.startsWith("Bearer ")) {
        return reply.status(401).send({ error: "Unauthorized: Missing or invalid authorization header" });
      }

      try {
        const token = authHeader.substring(7);
        
        // Log token info for debugging (remove in production)
        const tokenParts = token.split('.');
        if (tokenParts.length === 3) {
          const payload = JSON.parse(Buffer.from(tokenParts[1], 'base64').toString());
          fastify.log.info(`Token aud: ${payload.aud}, iss: ${payload.iss}, scp: ${payload.scp}`);
        }
        
        const { payload } = await jwtVerify(token, JWKS, {
          issuer: validIssuers,
          audience: validAudiences,
        });

        request.user = payload;
        request.token = token;
      } catch (err) {
        fastify.log.error(err, "JWT verification failed");
        return reply.status(401).send({ error: "Unauthorized: Invalid token" });
      }
    }
  );

  fastify.decorate("getToken", (request: FastifyRequest) => {
    return request.token;
  });
});
