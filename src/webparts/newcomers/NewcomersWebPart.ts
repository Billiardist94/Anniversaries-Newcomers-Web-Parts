import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import Newcomers from './components/Newcomers';
import { INewcomersProps } from './components/Newcomers';
import EmployeeDetailsService, { IEmployeeDetailsService } from '../../services/EmployeeDetailsService';
import { sp } from '@pnp/sp/presets/all';
import ComponentHelper from '../../helpers/ComponentHelper';
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';

export interface INewcomersWebPartProps {
  title: string;
  maxItems: number;
}

export default class NewcomersWebPart extends BaseClientSideWebPart<INewcomersWebPartProps> {
  private _service: IEmployeeDetailsService;
  private _card: any;

  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;
  
  public onInit(): Promise<void> {

    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._themeVariant = this._themeProvider.tryGetTheme();
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return ComponentHelper.loadSPComponentById('914330ee-2df2-4f6e-a858-30c23a812408').then((sharedLibrary: any) => {
      sp.setup({
        spfxContext: this.context as any
      });
      
      this._service = new EmployeeDetailsService(this.context, 'Employee Details');
      this._card = sharedLibrary.LivePersonaCard;
    });
  }

/**@param args */
private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
  console.log(args);
  this._themeVariant = args.theme;
  this.render();
}

  public render(): void {
    const element: React.ReactElement<INewcomersProps > = React.createElement(
      Newcomers,
      {
        maxItems: this.properties.maxItems,
        title: this.properties.title,
        displayMode: this.displayMode,
        service: this._service,
        onTitleChange: (value: string) =>  {
          this.properties.title = value;
        },
        card: this._card,
        serviceScope: this.context.serviceScope,
        themeVariant: this._themeVariant
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneSlider('maxItems', {
                  label: 'How many items to show?',
                  min: 1,
                  max: 20,
                  step: 1,
                  value: this.properties.maxItems
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
