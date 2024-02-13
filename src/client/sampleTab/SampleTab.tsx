import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, authentication } from "@microsoft/teams-js";
import jwtDecode from "jwt-decode";

/**
 * Implementation of the sample-tab content page
 */
export const SampleTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [tenantId, setTenantId] = useState<string | undefined>();
    const [groupId, setGroupId] = useState<string | undefined>();
    const [channelId, setChannelId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [preferredUser, setPreferredUser] = useState<String>();
    const [error, setError] = useState<string>();

    useEffect(() => {
        if (inTeams === true) {
            authentication.getAuthToken({
                resources: [process.env.TAB_APP_URI as string],
                silent: false
            } as authentication.AuthTokenRequestParameters).then(token => {
                const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
                setName(decoded!.name);
                setPreferredUser(decoded!.preferred_username)
                app.notifySuccess();
            }).catch(message => {
                setError(message);
                app.notifyFailure({
                    reason: app.FailedReason.AuthFailed,
                    message
                });
            });
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.page.id);
            setGroupId(context.team?.groupId)
            setChannelId(context.channel?.id);
        }
    }, [context]);

    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <Header content="This is your tab" />
                </Flex.Item>
                <Flex.Item>
                    <div>
                        <div>
                            <Text content={`Hello ${name}!`} />
                        </div>
                        <div>
                            <Text content={`Your preferred name is ${preferredUser}`} />
                        </div>
                        {error && <div><Text content={`An SSO error occurred ${error}`} /></div>}

                        <div>
                            Here is some more information we can retrieve from the Teams context
                        </div>
                        <div>
                            <Text content={`This channel ID ${channelId}!`} />
                        </div>
                        <div>
                            <Text content={`This group ID ${groupId}!`} />
                        </div>
                    </div>
                </Flex.Item>
                <Flex.Item styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Text size="smaller" content="(C) Copyright Microsoft" />
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
