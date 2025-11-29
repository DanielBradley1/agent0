import { streamText, tool } from "ai";
import ky from "ky";
import { z } from "zod";

import { openai } from "@ai-sdk/openai";

import { Message } from "./agent0-types.js";

// Exchange the API token for a Microsoft Graph token using OBO flow
async function getGraphToken(apiToken: string): Promise<string | null> {
  const tenantId = process.env.ENTRA_TENANT_ID;
  const clientId = process.env.ENTRA_CLIENT_ID;
  const clientSecret = process.env.ENTRA_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    console.error("Missing ENTRA_CLIENT_SECRET for OBO flow");
    return null;
  }

  try {
    const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    
    const params = new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      client_id: clientId,
      client_secret: clientSecret,
      assertion: apiToken,
      scope: "https://graph.microsoft.com/User.Read",
      requested_token_use: "on_behalf_of",
    });

    const response = await ky
      .post(tokenEndpoint, {
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: params.toString(),
      })
      .json<{ access_token: string }>();

    return response.access_token;
  } catch (error: unknown) {
    // Try to get the detailed error from the response
    if (error && typeof error === 'object' && 'response' in error) {
      const httpError = error as { response: Response };
      try {
        const errorBody = await httpError.response.json();
        console.error("OBO token exchange failed:", JSON.stringify(errorBody, null, 2));
      } catch {
        console.error("Error exchanging token via OBO flow:", error);
      }
    } else {
      console.error("Error exchanging token via OBO flow:", error);
    }
    return null;
  }
}

async function getUserInfo(apiToken: string) {
  try {
    // Exchange API token for Graph token using OBO flow
    const graphToken = await getGraphToken(apiToken);
    
    if (!graphToken) {
      return { error: "Failed to acquire Graph token. Ensure ENTRA_CLIENT_SECRET is configured." };
    }

    // Call Microsoft Graph API to get user information
    const response = await ky
      .get("https://graph.microsoft.com/v1.0/me", {
        headers: {
          Authorization: `Bearer ${graphToken}`,
        },
      })
      .json();
    return { result: JSON.stringify(response) };
  } catch (error) {
    console.error("Error fetching user info from Microsoft Graph:", error);
    return { error: "Failed to retrieve user information" };
  }
}

export async function agent0(messages: Message[], token: string) {
  const getUserInfoTool = tool({
    description: "Get information about the logged in user",
    parameters: z.object({}),
    execute: async () => await getUserInfo(token),
  });

  const stream = streamText({
    model: openai("gpt-4-turbo"),
    maxSteps: 5,
    tools: {
      getUserInfo: getUserInfoTool,
    },
    system: "assistant",
    messages: messages,
  });

  return stream;
}
