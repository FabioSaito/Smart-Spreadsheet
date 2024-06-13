require 'roo'

class TableDetector
  def self.call
    xlsx = Roo::Spreadsheet.open('tests/example_1.xlsx')
    sheet = xlsx.sheet(0)

    column_headers = get_column_headers(sheet)

    get_number_of_columns(sheet, column_headers)
  end

  def self.get_column_headers(sheet)
    tables_positions = []

    first_column = sheet.first_column
    last_column = sheet.last_column

    first_row = sheet.first_row
    last_row = sheet.last_row

    # loop through all cells to get the table positions
    (first_column..last_column).each do |column|

      row_start_position = 0
      number_of_rows = 0
      current_cell = sheet.cell(1, 1)

      (first_row..last_row + 2).each do |row|
        last_cell = current_cell
        current_cell = sheet.cell(row, column)

        row_start_position = row - 1 if (row_start_position == 0 && current_cell.is_a?(String) && last_cell.is_a?(String))
        number_of_rows += 1 if current_cell.is_a?(String)

        # add to tables_positions the found column header
        # zero the counter of rows
        if (number_of_rows > 1 && current_cell.nil? && last_cell.nil?)

          if sheet.cell(row_start_position + 1, column)[/\A\s*/].length > 0
            tables_positions << { row_start_position: row_start_position - 1, number_of_rows: number_of_rows, column: column, hierquical: true}
          else
            tables_positions << { row_start_position: row_start_position, number_of_rows: number_of_rows, column: column, hierquical: false}
          end

          row_start_position = 0
          number_of_rows = 0
        end
      end
    end

    tables_positions
  end

  def self.get_number_of_columns(sheet, column_headers)
    column_headers.each do |coord|
      row_start_position = coord[:row_start_position]
      number_of_rows = coord[:number_of_rows]
      column = coord[:column]
      number_of_columns = 0

      if coord[:hierquical]
        current_cell = sheet.cell(row_start_position + 1, column)
      else
        current_cell = sheet.cell(row_start_position, column)
      end

      (column..sheet.last_column + 2).each do |current_column|
        last_cell = current_cell
        current_cell = sheet.cell(row_start_position, current_column)
        number_of_columns += 1

        if current_cell.nil? && last_cell.nil?
          coord.merge!({number_of_columns: current_column - column - 1})
          break
        end
      end
    end
  end
end
