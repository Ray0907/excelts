/** Thrown for exceljs features this library deliberately does not implement; `feature` names what was used. */
export class NotSupportedError extends Error {
  constructor(
    readonly feature: string,
    detail = '',
  ) {
    super(`not supported: ${feature}${detail ? ` (${detail})` : ''}`);
    this.name = 'NotSupportedError';
  }
}
