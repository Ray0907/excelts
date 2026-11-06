export interface DataValidation {
  type?: string; operator?: string; allowBlank?: boolean; showInputMessage?: boolean; showErrorMessage?: boolean;
  formulae?: unknown[]; promptTitle?: string; prompt?: string; errorStyle?: string; errorTitle?: string; error?: string;
}

/** Validations keyed by cell address (a range key such as "A1:B2" is kept as the sqref). */
export class DataValidations {
  model: Record<string, DataValidation | null>;
  constructor(model?: Record<string, DataValidation | null>) {
    this.model = model ?? {};
  }
  add(address: string, validation: DataValidation | null): DataValidation | null {
    this.model[address] = validation;
    return validation;
  }
  find(address: string): DataValidation | null | undefined {
    return this.model[address];
  }
  remove(address: string): void {
    this.add(address, null);
  }
}
