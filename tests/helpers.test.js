// tests/helpers.test.js
import { findTextRange, getTableCellRange } from '../dist/googleDocsApiHelpers.js';
import assert from 'node:assert';
import { describe, it, mock } from 'node:test';

describe('Text Range Finding', () => {
  // Test hypothesis 1: Text range finding works correctly
  
  describe('findTextRange', () => {
    it('should find text within a single text run correctly', async () => {
      // Mock the docs.documents.get method to return a predefined structure
      const mockDocs = {
        documents: {
          get: mock.fn(async () => ({
            data: {
              body: {
                content: [
                  {
                    paragraph: {
                      elements: [
                        {
                          startIndex: 1,
                          endIndex: 25,
                          textRun: {
                            content: 'This is a test sentence.'
                          }
                        }
                      ]
                    }
                  }
                ]
              }
            }
          }))
        }
      };

      // Test finding "test" in the sample text
      const result = await findTextRange(mockDocs, 'doc123', 'test', 1);
      assert.deepStrictEqual(result, { startIndex: 11, endIndex: 15 });
      
      // Verify the docs.documents.get was called with the right parameters
      assert.strictEqual(mockDocs.documents.get.mock.calls.length, 1);
      assert.deepStrictEqual(
        mockDocs.documents.get.mock.calls[0].arguments[0], 
        {
          documentId: 'doc123',
          fields: 'body(content(paragraph(elements(startIndex,endIndex,textRun(content)))))'
        }
      );
    });
    
    it('should find the nth instance of text correctly', async () => {
      // Mock with a document that has repeated text
      const mockDocs = {
        documents: {
          get: mock.fn(async () => ({
            data: {
              body: {
                content: [
                  {
                    paragraph: {
                      elements: [
                        {
                          startIndex: 1,
                          endIndex: 41,
                          textRun: {
                            content: 'Test test test. This is a test sentence.'
                          }
                        }
                      ]
                    }
                  }
                ]
              }
            }
          }))
        }
      };

      // Find the 3rd instance of "test"
      const result = await findTextRange(mockDocs, 'doc123', 'test', 3);
      assert.deepStrictEqual(result, { startIndex: 27, endIndex: 31 });
    });

    it('should return null if text is not found', async () => {
      const mockDocs = {
        documents: {
          get: mock.fn(async () => ({
            data: {
              body: {
                content: [
                  {
                    paragraph: {
                      elements: [
                        {
                          startIndex: 1,
                          endIndex: 25,
                          textRun: {
                            content: 'This is a sample sentence.'
                          }
                        }
                      ]
                    }
                  }
                ]
              }
            }
          }))
        }
      };

      // Try to find text that doesn't exist
      const result = await findTextRange(mockDocs, 'doc123', 'test', 1);
      assert.strictEqual(result, null);
    });

    it('should handle text spanning multiple text runs', async () => {
      const mockDocs = {
        documents: {
          get: mock.fn(async () => ({
            data: {
              body: {
                content: [
                  {
                    paragraph: {
                      elements: [
                        {
                          startIndex: 1,
                          endIndex: 6,
                          textRun: {
                            content: 'This '
                          }
                        },
                        {
                          startIndex: 6,
                          endIndex: 11,
                          textRun: {
                            content: 'is a '
                          }
                        },
                        {
                          startIndex: 11,
                          endIndex: 20,
                          textRun: {
                            content: 'test case'
                          }
                        }
                      ]
                    }
                  }
                ]
              }
            }
          }))
        }
      };

      // Find text that spans runs: "a test"
      const result = await findTextRange(mockDocs, 'doc123', 'a test', 1);
      assert.deepStrictEqual(result, { startIndex: 9, endIndex: 15 });
    });
  });
});

describe('Table Cell Range Finding', () => {
  describe('getTableCellRange', () => {
    it('should find non-empty cell ranges correctly', async () => {
      // Mock a document with a 2x2 table containing text
      const mockDocs = {
        documents: {
          get: mock.fn(async () => ({
            data: {
              body: {
                content: [
                  {
                    startIndex: 10,
                    endIndex: 100,
                    table: {
                      tableRows: [
                        {
                          tableCells: [
                            {
                              startIndex: 11,
                              endIndex: 30,
                              content: [
                                {
                                  paragraph: {
                                    elements: [
                                      { startIndex: 12, endIndex: 25 }  // "Hello World\n"
                                    ]
                                  }
                                }
                              ]
                            },
                            {
                              startIndex: 31,
                              endIndex: 50,
                              content: [
                                {
                                  paragraph: {
                                    elements: [
                                      { startIndex: 32, endIndex: 40 }
                                    ]
                                  }
                                }
                              ]
                            }
                          ]
                        }
                      ]
                    }
                  }
                ]
              }
            }
          }))
        }
      };

      const result = await getTableCellRange(mockDocs, 'doc123', 10, 0, 0);
      assert.ok(result !== null);
      assert.strictEqual(result.cellStartIndex, 11);
      assert.strictEqual(result.cellEndIndex, 30);
      assert.strictEqual(result.contentStartIndex, 12);
      assert.strictEqual(result.contentEndIndex, 24); // 25 - 1 to exclude newline
      assert.strictEqual(result.paragraphEndIndex, 25); // Includes newline for paragraph styling
    });

    it('should return paragraphEndIndex > contentStartIndex for empty cells', async () => {
      // Mock a document with an empty cell (only contains newline)
      const mockDocs = {
        documents: {
          get: mock.fn(async () => ({
            data: {
              body: {
                content: [
                  {
                    startIndex: 10,
                    endIndex: 50,
                    table: {
                      tableRows: [
                        {
                          tableCells: [
                            {
                              startIndex: 11,
                              endIndex: 15,
                              content: [
                                {
                                  paragraph: {
                                    elements: [
                                      { startIndex: 12, endIndex: 13 }  // Just newline "\n"
                                    ]
                                  }
                                }
                              ]
                            }
                          ]
                        }
                      ]
                    }
                  }
                ]
              }
            }
          }))
        }
      };

      const result = await getTableCellRange(mockDocs, 'doc123', 10, 0, 0);
      assert.ok(result !== null);
      // For empty cells, contentStartIndex === contentEndIndex (no editable content)
      assert.strictEqual(result.contentStartIndex, 12);
      assert.strictEqual(result.contentEndIndex, 12);
      // But paragraphEndIndex includes the newline, enabling paragraph styling
      assert.strictEqual(result.paragraphEndIndex, 13);
      assert.ok(result.paragraphEndIndex > result.contentStartIndex,
        'paragraphEndIndex should include newline for paragraph styling on empty cells');
    });

    it('should throw UserError for out-of-bounds row index', async () => {
      const mockDocs = {
        documents: {
          get: mock.fn(async () => ({
            data: {
              body: {
                content: [
                  {
                    startIndex: 10,
                    endIndex: 50,
                    table: {
                      tableRows: [
                        { tableCells: [{ startIndex: 11, endIndex: 20 }] }
                      ]
                    }
                  }
                ]
              }
            }
          }))
        }
      };

      await assert.rejects(
        async () => await getTableCellRange(mockDocs, 'doc123', 10, 5, 0),
        (error) => {
          assert.ok(error.message.includes('Row index 5 out of bounds'));
          return true;
        }
      );
    });

    it('should throw UserError for out-of-bounds column index', async () => {
      const mockDocs = {
        documents: {
          get: mock.fn(async () => ({
            data: {
              body: {
                content: [
                  {
                    startIndex: 10,
                    endIndex: 50,
                    table: {
                      tableRows: [
                        { tableCells: [{ startIndex: 11, endIndex: 20 }] }
                      ]
                    }
                  }
                ]
              }
            }
          }))
        }
      };

      await assert.rejects(
        async () => await getTableCellRange(mockDocs, 'doc123', 10, 0, 5),
        (error) => {
          assert.ok(error.message.includes('Column index 5 out of bounds'));
          return true;
        }
      );
    });

    it('should throw UserError for invalid table index', async () => {
      const mockDocs = {
        documents: {
          get: mock.fn(async () => ({
            data: {
              body: {
                content: [
                  {
                    startIndex: 10,
                    endIndex: 50,
                    table: { tableRows: [] }
                  }
                ]
              }
            }
          }))
        }
      };

      await assert.rejects(
        async () => await getTableCellRange(mockDocs, 'doc123', 999, 0, 0),
        (error) => {
          assert.ok(error.message.includes('No table found at index 999'));
          return true;
        }
      );
    });
  });
});