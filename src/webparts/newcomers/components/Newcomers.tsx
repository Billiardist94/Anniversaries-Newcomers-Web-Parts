import * as React from 'react';
import styles from './Newcomers.module.scss';
import { IEmployeeDetailsService } from '../../../services/EmployeeDetailsService';
import UserProfile from '../../../models/UserProfile';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { Spinner, SpinnerSize, Icon, Label } from 'office-ui-fabric-react';
import { ServiceScope } from '@microsoft/sp-core-library';
import PersonaItem from '../../../components/persona/PersonaItem';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface INewcomersProps {
  maxItems: number;
  title: string;
  displayMode: number;
  service: IEmployeeDetailsService;
  onTitleChange(title: string): void;
  serviceScope: ServiceScope;
  card: any;
  themeVariant: IReadonlyTheme | undefined;
}

interface INewcomersState{
  employees: Array<UserProfile>;
  isLoading: boolean;
  error: string;
}

export default class Newcomers extends React.Component<INewcomersProps, INewcomersState> {
  constructor(props: INewcomersProps | Readonly<INewcomersProps>){
    super(props);

    this.state = {
      employees: new Array<UserProfile>(),
      isLoading: false,
      error: ''
    };
  }

  public componentDidMount(): void{
    this.setState({
      isLoading: true
    }, () => {
      this.props.service.getNewcomers(this.props.maxItems).then((employees: Array<UserProfile>) => {
        this.setState({
          employees: employees,
          isLoading: false
        });
      }).catch((ex) => {
        this.setState({
          error: ex.message,
          isLoading: false
        });
      });
    });
  }

  public render(): React.ReactElement<INewcomersProps> {
    const isError = this.state.error.length > 0;
    const isLoaded = !isError && !this.state.isLoading && this.state.employees.length > 0;
    const noItems = !isError && !this.state.isLoading && this.state.employees.length === 0;
    const { semanticColors }: IReadonlyTheme = this.props.themeVariant;
    const { palette }: IReadonlyTheme = this.props.themeVariant;

    return (
      <div className={ styles.newcomers }>
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.onTitleChange}
          themeVariant={this.props.themeVariant}
        />
        {
          this.state.isLoading &&
            <Spinner size={SpinnerSize.large} label={'Loading ...'} className={styles.spinner} />
        }
        {
          isError &&
            <div>{this.state.error}</div>
        }
        {
          noItems &&
          <div className={styles.noUsers}>
            <div className={styles.iconRow}>
              <Icon iconName={'ProfileSearch'} />
            </div>
            <div className={styles.row}>
              <Label>
                <span style={{color: semanticColors.bodyText}}>No newcomers found.</span>
              </Label>
            </div>
          </div>
        }
        {
          isLoaded &&
            this.state.employees.map((i) => {
              // eslint-disable-next-line react/jsx-key
              return <div className={styles.item} style={{color: this.props.themeVariant.semanticColors.bodyDivider}}>
                        <PersonaItem themeVariant={this.props.themeVariant} serviceScope={this.props.serviceScope} card={this.props.card} user={i} key={i.email} />
                        <a href={"mailto:"+ i.email}><Icon iconName="EditMail" className={styles.icon} style={{color: this.props.themeVariant.semanticColors.primaryButtonBackground, backgroundColor: this.props.themeVariant.semanticColors.bodyBackground}}/></a>
                    </div>;
            })
        }
      </div>
    );
  }
}
