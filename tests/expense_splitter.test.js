const { calculateExpensesWithSubgroups } = require('../expense_splitter');

// Mocking Google Apps Script objects
const SpreadsheetApp = {
    getActiveSpreadsheet: jest.fn(() => ({
        getActiveSheet: jest.fn(() => mockSheet)
    }))
};

global.SpreadsheetApp = SpreadsheetApp;

const Logger = {
    log: jest.fn()
};

global.Logger = Logger;

const mockSheet = {
    getDataRange: jest.fn(() => ({
        getValues: jest.fn(() => mockData)
    })),
    getRange: jest.fn(() => mockRange),
    getLastRow: jest.fn(() => 10),
    setRowHeights: jest.fn(),
    autoResizeColumns: jest.fn()
};

mockSheet.getDataRange.mockReturnValue({
    getValues: jest.fn(() => mockData)
});

const mockRange = {
    setValues: jest.fn(),
    setBorder: jest.fn(),
    setFontWeight: jest.fn(() => mockRange),
    setBackground: jest.fn(() => mockRange),
    setFontSize: jest.fn(() => mockRange),
    setHorizontalAlignment: jest.fn(() => mockRange),
    setFontColor: jest.fn(() => mockRange),
    getCell: jest.fn(() => ({ setBackground: jest.fn() })),
    getSheet: jest.fn(() => mockSheet),
    getNumRows: jest.fn(() => 5),
    getNumColumns: jest.fn(() => 5),
    offset: jest.fn(() => mockRange),
    getRow: jest.fn(() => 1),
    getColumn: jest.fn(() => 1)
};

const mockData = [
    ['Members', 'Subgroups', 'Item', 'Paid By', 'Amount', 'Split Between'],
    ['Alice', 'Group1', 'Lunch', 'Alice', '100', '*'],
    ['Bob', 'Group1', 'Dinner', 'Bob', '200', 'Alice,Bob']
];

describe('calculateExpensesWithSubgroups', () => {
    beforeEach(() => {
        jest.clearAllMocks();
    });

    it('should process expenses and write results to the sheet', () => {
        calculateExpensesWithSubgroups();

        // Verify data was read from the sheet
        expect(mockSheet.getDataRange).toHaveBeenCalled();
        expect(mockSheet.getDataRange().getValues).toHaveBeenCalled();

        // Verify results were written to the sheet
        expect(mockSheet.getRange).toHaveBeenCalled();
        expect(mockRange.setValues).toHaveBeenCalled();
    });

    it('should handle missing required columns gracefully', () => {
        const invalidData = [
            ['Invalid', 'Headers'],
            ['Data', 'Here']
        ];
        mockSheet.getDataRange.mockReturnValueOnce({
            getValues: jest.fn(() => invalidData)
        });

        calculateExpensesWithSubgroups();

        // Verify no data was written to the sheet
        expect(mockSheet.getRange).not.toHaveBeenCalled();
    });

    it('should calculate balances correctly', () => {
        calculateExpensesWithSubgroups();

        // Verify balances and splits are calculated
        expect(mockRange.setValues).toHaveBeenCalledWith(
            expect.arrayContaining([
                expect.arrayContaining(['Alice']),
                expect.arrayContaining(['Bob'])
            ])
        );
    });
});