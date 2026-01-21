// src/googleSlidesApiHelpers.ts
import { google, slides_v1 } from 'googleapis';
import { UserError } from 'fastmcp';

type Slides = slides_v1.Slides;

// --- Constants ---
const MAX_BATCH_UPDATE_REQUESTS = 50;
const EMU_PER_POINT = 12700; // 914400 EMU = 1 inch, 72 points = 1 inch

// --- EMU Conversion Helpers ---

/**
 * Converts points to EMU (English Metric Units)
 * Google Slides API uses EMU for all measurements
 * @param pt - Value in points
 * @returns Value in EMU
 */
export function emuFromPoints(pt: number): number {
  return Math.round(pt * EMU_PER_POINT);
}

/**
 * Converts EMU to points
 * @param emu - Value in EMU
 * @returns Value in points
 */
export function pointsFromEmu(emu: number): number {
  return emu / EMU_PER_POINT;
}

// --- Object ID Generation ---

/**
 * Generates a unique object ID for new elements
 * Format: prefix_timestamp_random
 * @param prefix - Optional prefix for the ID (default: 'obj')
 * @returns Unique object ID string
 */
export function generateObjectId(prefix: string = 'obj'): string {
  const timestamp = Date.now().toString(36);
  const random = Math.random().toString(36).substring(2, 8);
  return `${prefix}_${timestamp}_${random}`;
}

// --- Core Helper to Execute Batch Updates ---

/**
 * Executes a batch update on a presentation with error handling
 * @param slides - Google Slides API client
 * @param presentationId - The presentation ID
 * @param requests - Array of update requests
 * @returns Batch update response
 */
export async function executeBatchUpdate(
  slides: Slides,
  presentationId: string,
  requests: slides_v1.Schema$Request[]
): Promise<slides_v1.Schema$BatchUpdatePresentationResponse> {
  if (!requests || requests.length === 0) {
    return {};
  }

  if (requests.length > MAX_BATCH_UPDATE_REQUESTS) {
    console.warn(`Attempting batch update with ${requests.length} requests, exceeding typical limits. May fail.`);
  }

  try {
    const response = await slides.presentations.batchUpdate({
      presentationId: presentationId,
      requestBody: { requests },
    });
    return response.data;
  } catch (error: any) {
    console.error(`Google Slides API batchUpdate Error for presentation ${presentationId}:`, error.response?.data || error.message);

    if (error.code === 400) {
      const details = error.response?.data?.error?.details;
      let detailMsg = '';
      if (details && Array.isArray(details)) {
        detailMsg = details.map((d: any) => d.description || JSON.stringify(d)).join('; ');
      }
      throw new UserError(`Invalid request sent to Google Slides API. Details: ${detailMsg || error.message}`);
    }
    if (error.code === 404) {
      throw new UserError(`Presentation not found (ID: ${presentationId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied for presentation (ID: ${presentationId}). Ensure the authenticated user has edit access.`);
    }
    throw new Error(`Google Slides API Error (${error.code}): ${error.message}`);
  }
}

// --- Presentation Retrieval Helpers ---

/**
 * Gets full presentation data
 * @param slides - Google Slides API client
 * @param presentationId - The presentation ID
 * @returns Presentation data
 */
export async function getPresentation(
  slides: Slides,
  presentationId: string
): Promise<slides_v1.Schema$Presentation> {
  try {
    const response = await slides.presentations.get({
      presentationId: presentationId,
    });
    return response.data;
  } catch (error: any) {
    console.error(`Error getting presentation ${presentationId}:`, error.message);
    if (error.code === 404) {
      throw new UserError(`Presentation not found (ID: ${presentationId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied for presentation (ID: ${presentationId}). Ensure you have view access.`);
    }
    throw new Error(`Failed to get presentation: ${error.message}`);
  }
}

/**
 * Gets a specific slide/page from a presentation
 * @param slides - Google Slides API client
 * @param presentationId - The presentation ID
 * @param pageObjectId - The slide/page object ID
 * @returns Page data
 */
export async function getSlide(
  slides: Slides,
  presentationId: string,
  pageObjectId: string
): Promise<slides_v1.Schema$Page> {
  try {
    const response = await slides.presentations.pages.get({
      presentationId: presentationId,
      pageObjectId: pageObjectId,
    });
    return response.data;
  } catch (error: any) {
    console.error(`Error getting slide ${pageObjectId} from presentation ${presentationId}:`, error.message);
    if (error.code === 404) {
      throw new UserError(`Slide not found (ID: ${pageObjectId}) in presentation ${presentationId}. Check the IDs.`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied for presentation (ID: ${presentationId}).`);
    }
    throw new Error(`Failed to get slide: ${error.message}`);
  }
}

// --- Element Creation Helpers ---

/**
 * Creates a transform object for positioning elements
 * @param x - X position in points
 * @param y - Y position in points
 * @returns Transform object for API request
 */
export function createTransform(x: number, y: number): slides_v1.Schema$AffineTransform {
  return {
    scaleX: 1,
    scaleY: 1,
    shearX: 0,
    shearY: 0,
    translateX: emuFromPoints(x),
    translateY: emuFromPoints(y),
    unit: 'EMU',
  };
}

/**
 * Creates a size object for elements
 * @param width - Width in points
 * @param height - Height in points
 * @returns Size object for API request
 */
export function createSize(width: number, height: number): slides_v1.Schema$Size {
  return {
    width: { magnitude: emuFromPoints(width), unit: 'EMU' },
    height: { magnitude: emuFromPoints(height), unit: 'EMU' },
  };
}

/**
 * Creates a page element properties object
 * @param x - X position in points
 * @param y - Y position in points
 * @param width - Width in points
 * @param height - Height in points
 * @returns PageElementProperties object
 */
export function createPageElementProperties(
  x: number,
  y: number,
  width: number,
  height: number
): slides_v1.Schema$PageElementProperties {
  return {
    pageObjectId: undefined, // Will be set by caller
    size: createSize(width, height),
    transform: createTransform(x, y),
  };
}

/**
 * Converts hex color to RGB color object for Slides API
 * @param hex - Hex color string (e.g., '#FF0000' or 'FF0000')
 * @returns RgbColor object or null if invalid
 */
export function hexToRgbColor(hex: string): slides_v1.Schema$RgbColor | null {
  if (!hex) return null;
  let hexClean = hex.startsWith('#') ? hex.slice(1) : hex;

  if (hexClean.length === 3) {
    hexClean = hexClean[0] + hexClean[0] + hexClean[1] + hexClean[1] + hexClean[2] + hexClean[2];
  }
  if (hexClean.length !== 6) return null;

  const bigint = parseInt(hexClean, 16);
  if (isNaN(bigint)) return null;

  return {
    red: ((bigint >> 16) & 255) / 255,
    green: ((bigint >> 8) & 255) / 255,
    blue: (bigint & 255) / 255,
  };
}

// --- Request Builders ---

/**
 * Creates a request to add a new slide
 * @param insertionIndex - Optional index where to insert the slide
 * @param predefinedLayout - Optional layout type (BLANK, TITLE, TITLE_AND_BODY, etc.)
 * @param objectId - Optional object ID for the new slide
 * @returns CreateSlide request
 */
export function buildCreateSlideRequest(
  insertionIndex?: number,
  predefinedLayout?: string,
  objectId?: string
): slides_v1.Schema$Request {
  const request: slides_v1.Schema$Request = {
    createSlide: {
      objectId: objectId || generateObjectId('slide'),
      insertionIndex: insertionIndex,
      slideLayoutReference: predefinedLayout ? {
        predefinedLayout: predefinedLayout as slides_v1.Schema$LayoutReference['predefinedLayout'],
      } : undefined,
    },
  };
  return request;
}

/**
 * Creates a request to duplicate an object (slide or element)
 * @param objectId - The object to duplicate
 * @param objectIds - Map of old IDs to new IDs (optional)
 * @returns DuplicateObject request
 */
export function buildDuplicateObjectRequest(
  objectId: string,
  objectIds?: { [key: string]: string }
): slides_v1.Schema$Request {
  return {
    duplicateObject: {
      objectId: objectId,
      objectIds: objectIds,
    },
  };
}

/**
 * Creates a request to delete an object (slide or element)
 * @param objectId - The object to delete
 * @returns DeleteObject request
 */
export function buildDeleteObjectRequest(objectId: string): slides_v1.Schema$Request {
  return {
    deleteObject: {
      objectId: objectId,
    },
  };
}

/**
 * Creates a request to add a shape
 * @param pageObjectId - The slide to add the shape to
 * @param shapeType - Type of shape (TEXT_BOX, RECTANGLE, ELLIPSE, etc.)
 * @param x - X position in points
 * @param y - Y position in points
 * @param width - Width in points
 * @param height - Height in points
 * @param objectId - Optional object ID for the shape
 * @returns CreateShape request
 */
export function buildCreateShapeRequest(
  pageObjectId: string,
  shapeType: string,
  x: number,
  y: number,
  width: number,
  height: number,
  objectId?: string
): slides_v1.Schema$Request {
  return {
    createShape: {
      objectId: objectId || generateObjectId('shape'),
      shapeType: shapeType as slides_v1.Schema$CreateShapeRequest['shapeType'],
      elementProperties: {
        pageObjectId: pageObjectId,
        size: createSize(width, height),
        transform: createTransform(x, y),
      },
    },
  };
}

/**
 * Creates a request to add an image
 * @param pageObjectId - The slide to add the image to
 * @param imageUrl - URL of the image
 * @param x - X position in points
 * @param y - Y position in points
 * @param width - Width in points
 * @param height - Height in points
 * @param objectId - Optional object ID for the image
 * @returns CreateImage request
 */
export function buildCreateImageRequest(
  pageObjectId: string,
  imageUrl: string,
  x: number,
  y: number,
  width: number,
  height: number,
  objectId?: string
): slides_v1.Schema$Request {
  return {
    createImage: {
      objectId: objectId || generateObjectId('image'),
      url: imageUrl,
      elementProperties: {
        pageObjectId: pageObjectId,
        size: createSize(width, height),
        transform: createTransform(x, y),
      },
    },
  };
}

/**
 * Creates a request to add a table
 * @param pageObjectId - The slide to add the table to
 * @param rows - Number of rows
 * @param columns - Number of columns
 * @param x - X position in points
 * @param y - Y position in points
 * @param width - Width in points
 * @param height - Height in points
 * @param objectId - Optional object ID for the table
 * @returns CreateTable request
 */
export function buildCreateTableRequest(
  pageObjectId: string,
  rows: number,
  columns: number,
  x: number,
  y: number,
  width: number,
  height: number,
  objectId?: string
): slides_v1.Schema$Request {
  return {
    createTable: {
      objectId: objectId || generateObjectId('table'),
      rows: rows,
      columns: columns,
      elementProperties: {
        pageObjectId: pageObjectId,
        size: createSize(width, height),
        transform: createTransform(x, y),
      },
    },
  };
}

/**
 * Creates a request to insert text into a shape
 * @param objectId - The shape or table cell ID
 * @param text - The text to insert
 * @param insertionIndex - Optional index where to insert (0 for beginning)
 * @returns InsertText request
 */
export function buildInsertTextRequest(
  objectId: string,
  text: string,
  insertionIndex?: number
): slides_v1.Schema$Request {
  return {
    insertText: {
      objectId: objectId,
      text: text,
      insertionIndex: insertionIndex ?? 0,
    },
  };
}

/**
 * Creates a request to delete text from a shape
 * @param objectId - The shape or table cell ID
 * @param startIndex - Start index of text to delete
 * @param endIndex - End index of text to delete (exclusive)
 * @returns DeleteText request
 */
export function buildDeleteTextRequest(
  objectId: string,
  startIndex: number,
  endIndex: number
): slides_v1.Schema$Request {
  return {
    deleteText: {
      objectId: objectId,
      textRange: {
        type: 'FIXED_RANGE',
        startIndex: startIndex,
        endIndex: endIndex,
      },
    },
  };
}

/**
 * Creates a request to update slides position (reorder slides)
 * @param slideObjectIds - Array of slide IDs to move
 * @param insertionIndex - New position index
 * @returns UpdateSlidesPosition request
 */
export function buildUpdateSlidesPositionRequest(
  slideObjectIds: string[],
  insertionIndex: number
): slides_v1.Schema$Request {
  return {
    updateSlidesPosition: {
      slideObjectIds: slideObjectIds,
      insertionIndex: insertionIndex,
    },
  };
}

/**
 * Creates a request to update shape properties (fill color, etc.)
 * @param objectId - The shape ID
 * @param fillColor - Optional fill color in hex
 * @returns UpdateShapeProperties request or null if no changes
 */
export function buildUpdateShapePropertiesRequest(
  objectId: string,
  fillColor?: string
): slides_v1.Schema$Request | null {
  const shapeProperties: slides_v1.Schema$ShapeProperties = {};
  const fields: string[] = [];

  if (fillColor) {
    const rgbColor = hexToRgbColor(fillColor);
    if (rgbColor) {
      shapeProperties.shapeBackgroundFill = {
        solidFill: {
          color: { rgbColor },
        },
      };
      fields.push('shapeBackgroundFill.solidFill.color');
    }
  }

  if (fields.length === 0) return null;

  return {
    updateShapeProperties: {
      objectId: objectId,
      shapeProperties: shapeProperties,
      fields: fields.join(','),
    },
  };
}

/**
 * Creates a request to update text style (font size, color, bold, etc.)
 * @param objectId - The shape or text box ID
 * @param options - Style options to apply
 * @param options.fontSize - Font size in points (PT)
 * @param options.fontFamily - Font family name
 * @param options.bold - Whether text should be bold
 * @param options.italic - Whether text should be italic
 * @param options.foregroundColor - Text color in hex format
 * @param textRangeType - Type of text range ('ALL' for all text, or 'FIXED_RANGE')
 * @param startIndex - Start index for FIXED_RANGE
 * @param endIndex - End index for FIXED_RANGE
 * @returns UpdateTextStyle request or null if no style options provided
 */
export function buildUpdateTextStyleRequest(
  objectId: string,
  options: {
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    foregroundColor?: string;
  },
  textRangeType: 'ALL' | 'FIXED_RANGE' = 'ALL',
  startIndex?: number,
  endIndex?: number
): slides_v1.Schema$Request | null {
  const style: slides_v1.Schema$TextStyle = {};
  const fields: string[] = [];

  if (options.fontSize !== undefined) {
    style.fontSize = {
      magnitude: options.fontSize,
      unit: 'PT',
    };
    fields.push('fontSize');
  }

  if (options.fontFamily) {
    style.fontFamily = options.fontFamily;
    fields.push('fontFamily');
  }

  if (options.bold !== undefined) {
    style.bold = options.bold;
    fields.push('bold');
  }

  if (options.italic !== undefined) {
    style.italic = options.italic;
    fields.push('italic');
  }

  if (options.foregroundColor) {
    const rgbColor = hexToRgbColor(options.foregroundColor);
    if (rgbColor) {
      style.foregroundColor = {
        opaqueColor: {
          rgbColor: rgbColor,
        },
      };
      fields.push('foregroundColor');
    }
  }

  if (fields.length === 0) return null;

  const textRange: slides_v1.Schema$Range = {
    type: textRangeType,
  };

  if (textRangeType === 'FIXED_RANGE') {
    textRange.startIndex = startIndex;
    textRange.endIndex = endIndex;
  }

  return {
    updateTextStyle: {
      objectId: objectId,
      style: style,
      textRange: textRange,
      fields: fields.join(','),
    },
  };
}

// --- Speaker Notes Helpers ---

/**
 * Gets the speaker notes shape ID from a slide
 * @param slide - The slide page object
 * @returns The speaker notes shape ID or null
 */
export function getSpeakerNotesShapeId(slide: slides_v1.Schema$Page): string | null {
  const notesPage = slide.slideProperties?.notesPage;
  if (!notesPage?.pageElements) return null;

  // Find the shape with placeholder type BODY in the notes page
  for (const element of notesPage.pageElements) {
    if (element.shape?.placeholder?.type === 'BODY') {
      return element.objectId || null;
    }
  }
  return null;
}
