// Minimal PCF typings so TS can compile without pcf-scripts types.
// Supports StandardControl with 2 or 4 generics by providing defaults.

declare namespace ComponentFramework {
  interface Dictionary { [key: string]: any; }

  interface WebApi {
    createRecord(entityLogicalName: string, data: any): Promise<any>;
  }

  interface Utils {
    getEntityMetadata(logicalName: string): Promise<{ EntitySetName: string }>;
  }

  interface UserSettings { userId: string; }

  interface DatasetRecord {
    getFormattedValue(columnName: string): string | undefined;
    getValue(columnName: string): any;
  }

  interface Dataset {
    loading: boolean;
    records: { [id: string]: DatasetRecord };
    refresh(): Promise<void>;
    getTargetEntityType(): string;
  }

  interface Property { raw?: any; }

  // Inputs your control expects to see on context.parameters
  interface Parameters {
    dataset: Dataset;
    MessageProperty: Property;
    DateProperty: Property;
    UserProperty: Property;
    ParentLookUpProperty: Property;
  }

  interface Context<TInputs> {
    parameters: TInputs & Parameters & any;
    userSettings: UserSettings;
    mode: any;   // exposes contextInfo on forms
    utils: Utils;
    webAPI: WebApi;
  }

  // Provide defaults for 4 generics so using only <IInputs, IOutputs> still compiles.
  interface StandardControl<IInputs = any, IOutputs = any, TInputs = any, TOutputs = any> {
    init(
      context: Context<IInputs>,
      notifyOutputChanged: () => void,
      state: Dictionary,
      container: HTMLDivElement
    ): void;
    updateView(context: Context<IInputs>): void;
    getOutputs(): IOutputs;
    destroy(): void;
  }
}

