import * as React from 'react';
import UserProfile from '../../models/UserProfile';
import { ServiceScope } from '@microsoft/sp-core-library';
import { Persona, IPersonaStyles } from "office-ui-fabric-react";
import styles from './PersonaItem.module.scss';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IPersonaItem {
    card: any;
    user: UserProfile;
    serviceScope: ServiceScope;
    themeVariant: IReadonlyTheme | undefined;
}

export default class PersonaItem extends React.Component<IPersonaItem> {
    public render(): React.ReactElement {
        const personaStyles:IPersonaStyles = {
            root: {},
            details: {},
            optionalText: {},
            primaryText: {
                fontSize: "12px",
                color: this.props.themeVariant.semanticColors.bodySubtext,
                letterSpacing: "-0.16px",
                fontWeight: "600"
            },
            secondaryText: {
                color: this.props.themeVariant.semanticColors.bodySubtext,
            },
            tertiaryText: {},
            textContent: {}
        }; 
        let persona = <Persona styles={personaStyles} imageUrl={this.props.user.photoUrl} text={this.props.user.name} secondaryText={this.props.user.function} size={40} className={styles.item} />;
        if (this.props.card){
            persona = React.createElement(this.props.card, {
                className: 'people',
                clientScenario: 'PeopleWebPart',
                disableHover: false,
                hostAppPersonaInfo: {
                    PersonaType: 'User'
                },
                serviceScope: this.props.serviceScope,
                upn: this.props.user.email
            }, persona);
        }

        return persona;
    }
}