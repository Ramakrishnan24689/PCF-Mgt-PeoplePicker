import * as React from 'react';
import { useEffect, useState} from 'react';
import { InteractionRequiredAuthError, InteractionStatus } from "@azure/msal-browser";
import { AuthenticatedTemplate,UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { Login } from '@microsoft/mgt-react';

function ProtectedComponent() {

    const { instance, inProgress, accounts } = useMsal();
    const [apiData, setApiData] = useState(null);

    useEffect(() => {

        if (!apiData && inProgress === InteractionStatus.None) {
            const accessTokenRequest = {
                scopes: ["user.read"],
                account: accounts[0]
            }
         
            instance.acquireTokenSilent(accessTokenRequest).then((accessTokenResponse) => {
                // Acquire token silent success
                let accessToken = accessTokenResponse.accessToken;
                console.log(accessToken);
                // Call your API with token
               // callApi(accessToken).then((response:any) => { setApiData(response) });
            }).catch((error) => {
                if (error instanceof InteractionRequiredAuthError) {
                    instance.acquireTokenPopup(accessTokenRequest).then(function(accessTokenResponse) {
                        // Acquire token interactive success
                        let accessToken = accessTokenResponse.accessToken;
                        console.log(accessToken);
                        // Call your API with token
                      //  callApi(accessToken).then((response:any) => { setApiData(response) });
                    }).catch(function(error) {
                        // Acquire token interactive failure
                        console.log(error);
                    });
                }
                console.log(error);
            })
        }
    }, [instance, accounts, inProgress, apiData]);

    return <p>Return your protected content here: {apiData}</p>
}

function App() {
    return (
        <div>
        <Login/>
        <AuthenticatedTemplate>
            <ProtectedComponent />
        </ AuthenticatedTemplate>
        </div>
    )
}

export const MGTPeoplePicker: React.FunctionComponent = props => {
   
    
    return (
        <div>
            <App/>
        </div>
    );
}