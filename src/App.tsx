import { useEffect, useState } from "react";
import './App.css';
import { AccountInfo, IPublicClientApplication, PublicClientNext, AuthenticationResult, AuthError, InteractionRequiredAuthError, ServerError } from "@azure/msal-browser";

const clientId = "50a25558-6bab-41c6-82d2-6f76bc4ebd34";
const config = {
  auth: {
    clientId,
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin,
    supportsNestedAppAuth: true,
    // clientCapabilities: ["CP1"],
  },
  cache: {
    cacheLocation: "localStorage",
  }
};

let activeAccount: AccountInfo | null = null;
let msalInstanceInternal: IPublicClientApplication | null = null;
async function ensureMsalInitialized(): Promise<IPublicClientApplication> {
  if (msalInstanceInternal == null) {
    msalInstanceInternal = await PublicClientNext.createPublicClientApplication(config);
    activeAccount = msalInstanceInternal.getActiveAccount();
  }
  return msalInstanceInternal;
}

async function initializeOfficeJs() {
  let result = await Office.onReady();
  console.log(`Initialized with ${result}`);
}

function App() {
  const [scope, setScope] = useState('user.read');
  const [claims, setClaims] = useState('');
  const [isPopup, setIsPopup] = useState(false);
  const [output, setOutput] = useState<any>(null);
  const [timestamp, setTimestamp] = useState("");
  const [accessToken, setAccessToken] = useState("");

  useEffect(() => {
    initializeOfficeJs();
  }, []);

  const acquireTokenMsalJs = async () => {
    const startTime = new Date();
    const msalInstance = await ensureMsalInitialized();
    let loginHint = ""; // hardcode loginhint here to work in OWA
    try {
      loginHint = Office.context.mailbox.userProfile.emailAddress;
    } catch {}
  
    const requestParams = {
      scopes: scope.split(' '),
      loginHint: loginHint,
      ...(claims && { claims })
    }
    let result: AuthenticationResult | null = null;
    try {
      if (isPopup) {
        result = await msalInstance.acquireTokenPopup({
          ...requestParams
        });
      } else if (activeAccount) {
        result = await msalInstance.acquireTokenSilent({
          ...requestParams,
          account: activeAccount,
        });
      } else {
        result = await msalInstance.ssoSilent({
          ...requestParams,
        });
      }
      activeAccount = result.account;
      setOutput(result);
      setAccessToken(result.accessToken);
      msalInstance.setActiveAccount(result.account);
    } catch (ex) {
      let authError = ex as AuthError;
      let authErrorType = "unknown";
      if (authError instanceof InteractionRequiredAuthError) {
        authErrorType = "interaction_required";
      } else if (authError instanceof ServerError) {
        authErrorType = "server_error";
      }
      console.log(ex);
      setOutput({
        error: authError,
        code: authError.errorCode,
        message: authError.errorMessage,
        subError: authError.subError,
        authErrorType
      });
    }
    const endTime = new Date();
    setTimestamp(`${endTime.getTime() - startTime.getTime()}ms`);
  };
  const acquireTokenOfficeJs = async () => {
    const startTime = new Date();
    // Make sure office.js is ready
    try {
      await Office.onReady();
      const response = await (window as any).OfficeFirstPartyAuth.NestedAppAuth.getAccessToken({clientId, scope, silent: !isPopup, claims});
      console.log(response);
      setOutput(response);
    } catch (ex) {
      setOutput(ex);
      console.log(ex);
    }
    const endTime = new Date();
    setTimestamp(`${endTime.getTime() - startTime.getTime()}ms`);
  };
  const clearOutput = () => {
    setOutput(null);
    setTimestamp("");
  };
  const prepopulateClaims = () => {
    const claims = {"access_token":{"nbf":{"essential":true,"value":Math.floor(new Date().getTime() / 1000 - 300).toString()}}};
    setClaims(JSON.stringify(claims));
  };
  const makeGraphCall = async (endpointUrl: string) => {
    let files = await fetch(endpointUrl, {
      method: 'GET',
      headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
      }
    });
    let responseJson = await files.json();
    setOutput(responseJson);
  }

  const filesGraphCall = () => {
    // const endpointUrl = 'https://graph.microsoft.com/v1.0/sites/root/drive/items/01DTQ7UIHXH3DWPZYZ6JGZUED7GNP4LOGP';
    // const endpointUrl = 'https://graph.microsoft.com/v1.0/sites/18d5d58e-cd4f-4ad4-ac49-692b7cf47bd3/drive/root/children';
    // Home Tenant
    const endpointUrl = 'https://graph.microsoft.com/v1.0/sites/root/drive/items/01SOMZBKOMKSIBHJMXM5GLBP2BFRM5WS6B/children';
    makeGraphCall(endpointUrl);
  };
  const profileGraphCall = () => {
    makeGraphCall("https://graph.microsoft.com/v1.0/me");
  };

  return (
    <div className="App">
      <div className="App-content">
        <h1>NAA Test App</h1>
        <div className="input-group">
            <label htmlFor="scope-input" className="input-label">Scope:</label>
            <input
              type="text"
              id="scope-input"
              placeholder="Enter scope"
              value={scope}
              onChange={(e) => setScope(e.target.value)}
            />
        </div>
        <div className="input-group">
            <label htmlFor="claims-input" className="input-label">Claims:</label>
            <input
              type="text"
              id="claims-input"
              placeholder="Enter claims"
              value={claims}
              onChange={(e) => setClaims(e.target.value)}
            />
            <button onClick={prepopulateClaims}>Pre populate</button>
        </div>
        <div className="input-group checkbox-group">
            <label htmlFor="isPopup" className="input-label">Is Popup:</label>
            <input
              type="checkbox"
              id="isPopup"
              checked={isPopup}
              onChange={(e) => setIsPopup(e.target.checked)}
            />
        </div>
        <button onClick={acquireTokenOfficeJs}>Acquire Token Office.js</button>
        <button onClick={acquireTokenMsalJs}>Acquire Token MSAL.js</button>
        <button onClick={clearOutput}>Clear output</button>
        <button onClick={filesGraphCall}>Files Graph Call</button>
        <button onClick={profileGraphCall}>Profile Graph Call</button>
        <button onClick={() => window.location.reload()} className="reload-button">Reload</button>
        {timestamp && (<p>Time for request: {timestamp}</p>)}
        {output && (
          <div className="token-info">
            <h2>Token Information:</h2>
            <pre>{JSON.stringify(output, null, 2)}</pre>  {/* Formatting JSON data */}
          </div>
        )}
        <div>Running at: {window.location.origin}</div>
      </div>
    </div>
  );
}

export default App;
