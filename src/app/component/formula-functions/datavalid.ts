export class DataValidator {
  static readonly DATE_REGEX = /^\d{4}-\d{2}-\d{2}$/;
  static readonly NUMBER_REGEX = /^-?\d*\.?\d*$/;

  static validateCell(value: string, type: 'text' | 'number' | 'date'): { isValid: boolean; message?: string } {
    if (!value) return { isValid: true };

    switch (type) {
      case 'number':
        return this.validateNumber(value);
      case 'date':
        return this.validateDate(value);
      case 'text':
        return { isValid: true };
      default:
        return { isValid: true };
    }
  }

  private static validateNumber(value: string): { isValid: boolean; message?: string } {
    if (!this.NUMBER_REGEX.test(value)) {
      return {
        isValid: false,
        message: 'Please enter a valid number'
      };
    }
    return { isValid: true };
  }

  private static validateDate(value: string): { isValid: boolean; message?: string } {
    if (!this.DATE_REGEX.test(value)) {
      return {
        isValid: false,
        message: 'Please enter a date in YYYY-MM-DD format'
      };
    }
    const date = new Date(value);
    if (isNaN(date.getTime())) {
      return {
        isValid: false,
        message: 'Please enter a valid date'
      };
    }
    return { isValid: true };
  }
}
