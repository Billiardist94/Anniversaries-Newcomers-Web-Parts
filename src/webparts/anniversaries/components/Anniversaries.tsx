import * as React from 'react';
import styles from './Anniversaries.module.scss';
import { IEmployeeDetailsService } from '../../../services/EmployeeDetailsService';
import { ServiceScope } from '@microsoft/sp-core-library';
import UserProfile from '../../../models/UserProfile';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { Spinner, SpinnerSize, Icon, Label, Link } from 'office-ui-fabric-react';
import PersonaItem from '../../../components/persona/PersonaItem';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

interface IAnniversariesState{
  employees: Array<UserProfile>;
  isLoading: boolean;
  error: string;
}

export interface IAnniversariesProps {
  maxItems: number;
  title: string;
  displayMode: number;
  service: IEmployeeDetailsService;
  onTitleChange(title: string): void;
  serviceScope: ServiceScope;
  card: any;
  range: string;
  moreLink: string;
  themeVariant: IReadonlyTheme | undefined;
}

export default class Anniversaries extends React.Component<IAnniversariesProps, IAnniversariesState> {
  constructor(props: IAnniversariesProps | Readonly<IAnniversariesProps>){
    super(props);

    this.state = {
      employees: new Array<UserProfile>(),
      isLoading: false,
      error: ''
    }
  }

  public componentDidMount(): void{
    this.setState({
      isLoading: true
    }, () => {
      this.props.service.getAnniversaries(this.props.maxItems, this.props.range).then((employees: Array<UserProfile>) => {
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

  public render(): React.ReactElement<IAnniversariesProps> {
    const isError = this.state.error.length > 0;
    const isLoaded = !isError && !this.state.isLoading && this.state.employees.length > 0;
    const noItems = !isError && !this.state.isLoading && this.state.employees.length === 0;
    const { palette }: IReadonlyTheme = this.props.themeVariant;
    const { semanticColors }: IReadonlyTheme = this.props.themeVariant;

    const moreLink = this.props.moreLink && this.props.moreLink.length > 0 && this.state.employees.length > 0 ? <Link href={this.props.moreLink}>See all</Link> : <span></span>;
    
    return (
      <div className={ styles.anniversaries }>
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.onTitleChange}
          moreLink={moreLink}
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
                <span style={{color: semanticColors.bodyText}}>No anniversaries found at this.</span>
              </Label>
            </div>
          </div>
        }
        {
          isLoaded &&
            <div className={styles.items}>
              {
                this.state.employees.map((i) => {
                  const years = (new Date()).getFullYear() - i.hireDate.getFullYear();
                  if (years === 0){
                    return <div></div>;
                  }
    
                  // eslint-disable-next-line react/jsx-key
                  return <div className={styles.item} style={{color: this.props.themeVariant.semanticColors.bodyDivider}}>
                            <PersonaItem themeVariant={this.props.themeVariant} serviceScope={this.props.serviceScope} card={this.props.card} user={i} key={i.email} />
                            
                            <div className={styles.anniv}>
                              <img src={
                                years < 3 
                                ? require('../assets/bronze.png') 
                                : years >=3 && years < 8
                                ? require('../assets/silver.png')
                                : years >=8
                                ? require('../assets/gold.png')
                                : true
                                } alt="medal" className={styles.medal} 
                              />
                              <span className={styles.number}>{years}</span>
                              <span className={styles.text}>year{ years > 1 ? 's' : ''}</span>
                            </div>
                        </div>;
                })
              }
            </div>
        }
      </div>
    );
  }
}
