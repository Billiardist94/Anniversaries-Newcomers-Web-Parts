import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';

import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';
import ComponentHelper from '../../helpers/ComponentHelper';
import * as strings from 'AnniversariesWebPartStrings';
import Anniversaries, { IAnniversariesProps } from './components/Anniversaries';

import EmployeeDetailsService, { IEmployeeDetailsService } from '../../services/EmployeeDetailsService';
import { sp } from '@pnp/sp/presets/all';

export interface IAnniversariesWebPartProps {
  maxItems: number;
  title: string;
  range: string;
  moreLink: string;
}

export default class AnniversariesWebPart extends BaseClientSideWebPart<IAnniversariesWebPartProps> {

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
    this._themeVariant = args.theme;
    this.render();
  }

  public render(): void {
    const element: React.ReactElement<IAnniversariesProps> = React.createElement(
      Anniversaries,
      {
        maxItems: this.properties.maxItems,
        title: this.properties.title,
        service: this._service,
        card: this._card,
        serviceScope: this.context.serviceScope,
        displayMode: this.displayMode,
        onTitleChange: (value: string) => {
          this.properties.title = value;
        },
        range: this.properties.range,
        moreLink: this.properties.moreLink,
        themeVariant: this._themeVariant
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneSlider('maxItems', {
                  label: 'How many items to show? 0 - All',
                  min: 0,
                  max: 20,
                  step: 1,
                  value: this.properties.maxItems
                }),
                PropertyPaneDropdown('range', {
                  label: 'Date Range',
                  selectedKey: this.properties.range,
                  options: [
                    {
                      key: 'Day',
                      text: 'This Day'
                    },
                    {
                      key: 'Week',
                      text: 'This Week'
                    },
                    {
                      key: 'Month',
                      text: 'This Month'
                    }
                  ]
                }),
                PropertyPaneTextField('moreLink', {
                  label: 'More items links',
                  value: this.properties.moreLink
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
