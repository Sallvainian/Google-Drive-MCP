// tests/slides.test.js
import {
  emuFromPoints,
  pointsFromEmu,
  generateObjectId,
  executeBatchUpdate,
  hexToRgbColor,
  buildCreateSlideRequest,
  buildCreateShapeRequest,
  buildDeleteObjectRequest,
} from '../dist/googleSlidesApiHelpers.js';
import assert from 'node:assert';
import { describe, it, mock } from 'node:test';

describe('EMU Conversion', () => {
  describe('emuFromPoints', () => {
    it('should convert 1 point to 12700 EMU', () => {
      const result = emuFromPoints(1);
      assert.strictEqual(result, 12700);
    });

    it('should convert 72 points (1 inch) to 914400 EMU', () => {
      const result = emuFromPoints(72);
      assert.strictEqual(result, 914400);
    });

    it('should round fractional EMU values', () => {
      const result = emuFromPoints(0.5);
      assert.strictEqual(result, 6350);
    });

    it('should handle 0 points', () => {
      const result = emuFromPoints(0);
      assert.strictEqual(result, 0);
    });
  });

  describe('pointsFromEmu', () => {
    it('should convert 12700 EMU to 1 point', () => {
      const result = pointsFromEmu(12700);
      assert.strictEqual(result, 1);
    });

    it('should convert 914400 EMU (1 inch) to 72 points', () => {
      const result = pointsFromEmu(914400);
      assert.strictEqual(result, 72);
    });

    it('should handle 0 EMU', () => {
      const result = pointsFromEmu(0);
      assert.strictEqual(result, 0);
    });

    it('should preserve precision for fractional values', () => {
      const result = pointsFromEmu(6350);
      assert.strictEqual(result, 0.5);
    });
  });

  describe('round-trip conversion', () => {
    it('should be reversible for integer points', () => {
      const original = 100;
      const emu = emuFromPoints(original);
      const backToPoints = pointsFromEmu(emu);
      assert.strictEqual(backToPoints, original);
    });
  });
});

describe('Object ID Generation', () => {
  describe('generateObjectId', () => {
    it('should generate unique IDs', () => {
      const id1 = generateObjectId();
      const id2 = generateObjectId();
      assert.notStrictEqual(id1, id2);
    });

    it('should use default prefix "obj"', () => {
      const id = generateObjectId();
      assert.ok(id.startsWith('obj_'));
    });

    it('should use custom prefix when provided', () => {
      const id = generateObjectId('slide');
      assert.ok(id.startsWith('slide_'));
    });

    it('should generate IDs with consistent format', () => {
      const id = generateObjectId('test');
      const parts = id.split('_');
      assert.strictEqual(parts.length, 3); // prefix_timestamp_random
      assert.strictEqual(parts[0], 'test');
    });
  });
});

describe('Hex to RGB Color Conversion', () => {
  describe('hexToRgbColor', () => {
    it('should convert 6-digit hex to RGB (with #)', () => {
      const result = hexToRgbColor('#FF0000');
      assert.deepStrictEqual(result, { red: 1, green: 0, blue: 0 });
    });

    it('should convert 6-digit hex to RGB (without #)', () => {
      const result = hexToRgbColor('00FF00');
      assert.deepStrictEqual(result, { red: 0, green: 1, blue: 0 });
    });

    it('should convert 3-digit hex to RGB', () => {
      const result = hexToRgbColor('#F00');
      assert.deepStrictEqual(result, { red: 1, green: 0, blue: 0 });
    });

    it('should handle white color', () => {
      const result = hexToRgbColor('#FFFFFF');
      assert.deepStrictEqual(result, { red: 1, green: 1, blue: 1 });
    });

    it('should handle black color', () => {
      const result = hexToRgbColor('#000000');
      assert.deepStrictEqual(result, { red: 0, green: 0, blue: 0 });
    });

    it('should return null for empty string', () => {
      const result = hexToRgbColor('');
      assert.strictEqual(result, null);
    });

    it('should return null for invalid hex', () => {
      const result = hexToRgbColor('invalid');
      assert.strictEqual(result, null);
    });
  });
});

describe('Request Builders', () => {
  describe('buildCreateSlideRequest', () => {
    it('should build a basic create slide request', () => {
      const request = buildCreateSlideRequest(undefined, undefined, 'test_id');
      assert.ok(request.createSlide);
      assert.strictEqual(request.createSlide.objectId, 'test_id');
    });

    it('should include insertion index when provided', () => {
      const request = buildCreateSlideRequest(2, undefined, 'test_id');
      assert.strictEqual(request.createSlide.insertionIndex, 2);
    });

    it('should include layout when provided', () => {
      const request = buildCreateSlideRequest(undefined, 'TITLE', 'test_id');
      assert.deepStrictEqual(request.createSlide.slideLayoutReference, {
        predefinedLayout: 'TITLE',
      });
    });
  });

  describe('buildCreateShapeRequest', () => {
    it('should build a shape request with correct EMU values', () => {
      const request = buildCreateShapeRequest(
        'page_id',
        'RECTANGLE',
        100, // x in points
        50, // y in points
        200, // width in points
        100, // height in points
        'shape_id'
      );

      assert.ok(request.createShape);
      assert.strictEqual(request.createShape.shapeType, 'RECTANGLE');
      assert.strictEqual(request.createShape.objectId, 'shape_id');
      assert.strictEqual(request.createShape.elementProperties.pageObjectId, 'page_id');

      // Verify EMU conversion
      const size = request.createShape.elementProperties.size;
      assert.strictEqual(size.width.magnitude, 200 * 12700); // 200 points in EMU
      assert.strictEqual(size.height.magnitude, 100 * 12700); // 100 points in EMU
    });
  });

  describe('buildDeleteObjectRequest', () => {
    it('should build a delete object request', () => {
      const request = buildDeleteObjectRequest('element_id');
      assert.deepStrictEqual(request, {
        deleteObject: {
          objectId: 'element_id',
        },
      });
    });
  });
});

describe('Batch Update Execution', () => {
  describe('executeBatchUpdate', () => {
    it('should return empty object for empty requests array', async () => {
      const mockSlides = {};
      const result = await executeBatchUpdate(mockSlides, 'pres_id', []);
      assert.deepStrictEqual(result, {});
    });

    it('should call presentations.batchUpdate with correct parameters', async () => {
      const mockSlides = {
        presentations: {
          batchUpdate: mock.fn(async () => ({
            data: { replies: [{ createSlide: { objectId: 'new_slide_id' } }] },
          })),
        },
      };

      const requests = [{ createSlide: { objectId: 'new_slide_id' } }];
      await executeBatchUpdate(mockSlides, 'pres_id', requests);

      assert.strictEqual(mockSlides.presentations.batchUpdate.mock.calls.length, 1);
      const callArgs = mockSlides.presentations.batchUpdate.mock.calls[0].arguments[0];
      assert.strictEqual(callArgs.presentationId, 'pres_id');
      assert.deepStrictEqual(callArgs.requestBody.requests, requests);
    });

    it('should throw UserError for 404 response', async () => {
      const mockSlides = {
        presentations: {
          batchUpdate: mock.fn(async () => {
            const error = new Error('Not found');
            error.code = 404;
            throw error;
          }),
        },
      };

      await assert.rejects(
        executeBatchUpdate(mockSlides, 'invalid_id', [{ deleteObject: {} }]),
        (error) => {
          return error.message.includes('Presentation not found');
        }
      );
    });

    it('should throw UserError for 403 response', async () => {
      const mockSlides = {
        presentations: {
          batchUpdate: mock.fn(async () => {
            const error = new Error('Forbidden');
            error.code = 403;
            throw error;
          }),
        },
      };

      await assert.rejects(
        executeBatchUpdate(mockSlides, 'pres_id', [{ deleteObject: {} }]),
        (error) => {
          return error.message.includes('Permission denied');
        }
      );
    });
  });
});
