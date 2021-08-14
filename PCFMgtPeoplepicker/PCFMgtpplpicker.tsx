import * as React from 'react';
import { useState, useEffect } from 'react';
import { PeoplePicker, People, PersonType, } from '@microsoft/mgt-react';
import { Providers, ProviderState, SimpleProvider } from '@microsoft/mgt-element';
import { PublicClientApplication } from "@azure/msal-browser";
import { IInputs } from "./generated/ManifestTypes";
import { Spinner } from '@fluentui/react/lib/Spinner';


export interface IPeopleProps {
    people?: any;
    preselectedPeople?: any;
    context?: ComponentFramework.Context<IInputs>;
    peopleList?: (newValue: any) => void;
    isPickerDisabled?: boolean;
    toShowSelectedPeople?: boolean;
    pickerType?: string | number;
    selectionMode?: string;
    redirectUri?: string;
    clientId?: string;
    authority?: string;
}

export interface IPeoplePersona {
    id?: string;
    userPrincipalName?: string;
    displayName?: string;
}


export const MGTPeoplePicker: React.FunctionComponent<IPeopleProps> = props => {

    const [gotAccessToken] = useGotAccessToken();
    const [defaultPickerValue, setDefaultPickerValue] = useState<string[]>([]);
    const [people, setPeople] = useState([]);
    const msalConfig = {
        auth: {
            clientId: props.clientId!,
            authority: props.authority,
            redirectUri: props.redirectUri,
            postLogoutRedirectUri: window.location.href
        },
        cache: {
            cacheLocation: "sessionStorage", // This configures where your cache will be stored
            storeAuthStateInCookie: true, // Set this to "true" if you are having issues on IE11 or Edge
        }
    };
    const publicClientApplication = new PublicClientApplication(msalConfig);
    let provider: SimpleProvider;
    // Add scopes here for ID token to be used at Microsoft identity platform endpoints.
    const ssoRequest: any = {
        scopes: ["User.Read", "People.Read", "User.Read.All", "Group.Read.All", "User.ReadBasic.All"]
    };
    const handleSelectionChanged = async (e: any) => {
        let tempPeopleList: IPeoplePersona[] = [];
        await Promise.all(e.target.selectedPeople.map((currentPerson: IPeoplePersona) => {
            if (currentPerson)
                tempPeopleList.push({ id: currentPerson.id, userPrincipalName: currentPerson.userPrincipalName, displayName: currentPerson.displayName });
        }));
        props.peopleList!(tempPeopleList);
        setPeople(e.target.selectedPeople);
    };

    function getCurrentUserEmailID(): Promise<string> {
        return new Promise(async (resolve: any, reject: any) => {
            try {
                const _Users = await props!.context!.webAPI.retrieveRecord("systemuser", props.context?.userSettings.userId!);
                resolve(_Users.internalemailaddress);
            } catch (err) { console.log(err); reject(err); }
        });
    }

    function getAccessTokenByMSAL(userEmailID: string): Promise<string> {
        return new Promise(async (resolve: any, reject: any) => {
            try {
                const publicClientApplication = new PublicClientApplication(msalConfig);
                ssoRequest.loginHint = userEmailID;
                const ssoResponse = await publicClientApplication.ssoSilent(ssoRequest);
                resolve(ssoResponse.accessToken);
            } catch (err) { console.log(err); reject(err); }
        });
    }


    function useGotAccessToken(): [boolean] {
        const [gotAccessToken, setGotAccessToken] = useState(false);
        let tempDefaultPeople: string[] = [];
        useEffect(() => {
            const updateState = async () => {
                if (provider === undefined) {
                    const emailID: string = await getCurrentUserEmailID();
                    const accessToken: string = await getAccessTokenByMSAL(emailID);

                    sessionStorage.setItem("webApiAccessToken", accessToken);
                    provider = new SimpleProvider(async function getAccessTokenhandler(scopes: string[]) {
                        try {
                            let _accessToken = sessionStorage.getItem("webApiAccessToken");
                            if (_accessToken) {
                                return _accessToken;
                            }
                            else {
                                ssoRequest.loginHint = emailID;
                                let response = await publicClientApplication.ssoSilent(ssoRequest);
                                _accessToken = response.accessToken;
                                sessionStorage.setItem("webApiAccessToken", _accessToken);
                                return _accessToken;
                            }
                        } catch (error) {
                            console.log(error);
                            // see if pop up interaction to aquire token is required
                            return error;
                        }
                    });
                    Providers.globalProvider = provider;
                    Providers.globalProvider.setState(ProviderState.SignedIn);

                }
                setGotAccessToken(true);

                Promise.all(props.preselectedPeople.map((currentPerson: IPeoplePersona) => {
                    if (currentPerson)
                        tempDefaultPeople.push(currentPerson.id!);
                }));
                setDefaultPickerValue(tempDefaultPeople);
            };

            Providers.onProviderUpdated(updateState);
            updateState();

            return () => {
                Providers.removeProviderUpdatedListener(updateState);
            }
        }, []);

        return [gotAccessToken];
    }

    return (
        <div>
            {gotAccessToken && <div><PeoplePicker {...props.pickerType === "person" ? { type: PersonType.person } :
                { ...props.pickerType === "group" ? { type: PersonType.group } : { type: PersonType.any } }} {...props.isPickerDisabled ? { disabled: true } : undefined}
                defaultSelectedUserIds={defaultPickerValue} selectionMode={props.selectionMode} selectionChanged={handleSelectionChanged} />
                {props.toShowSelectedPeople && <div>Selected People: <People people={people} /></div>}</div>}
            {!gotAccessToken && <Spinner label="Loading..." ariaLive="assertive" labelPosition="left" />}
        </div>
    );
}

