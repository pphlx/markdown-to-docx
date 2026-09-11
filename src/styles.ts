import { BorderStyle, AlignmentType, WidthType } from 'docx';

export interface ConverterOptions {
  headerImage?: string;
  footerImage?: string;
  watermarkText?: string;
  accentColor?: string;
  openAfterConversion?: boolean;
}

export const DEFAULT_ACCENT_COLOR = 'B48C3C'; // Gold / Bronze accent
export const DEFAULT_DARK_COLOR = '0F172A';
export const DEFAULT_LIGHT_BG = 'F8FAFC';
export const DEFAULT_TEXT_COLOR = '1E293B';

export const TABLE_BORDER_STYLE = {
  style: BorderStyle.SINGLE,
  size: 4,
  color: 'CBD5E1',
};

export const CELL_MARGINS = {
  top: 100,
  bottom: 100,
  left: 140,
  right: 140,
};
