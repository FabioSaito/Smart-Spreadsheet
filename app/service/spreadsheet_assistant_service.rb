class SpreadsheetAssistantService
  def initialize(question:)
    @question = question
  end

  def call
    AiAssistantService.call(parsed_spreadsheet, @question)
  end

  def parsed_spreadsheet
    xlsx = Roo::Spreadsheet.open('tests/example_0.xlsx')
    sheet = xlsx.sheet(0)

    simple_table_parameters = [
      { row_start_position: 1, number_of_rows: 15 },
      { row_start_position: 19, number_of_rows: 17 },
      { row_start_position: 39, number_of_rows: 9 },
      { row_start_position: 51, number_of_rows: 8 },
      { row_start_position: 62, number_of_rows: 2 },
      { row_start_position: 67, number_of_rows: 3 },
      { row_start_position: 73, number_of_rows: 8 }
    ]

    spreadsheet_informations = process_simple_tables(sheet, simple_table_parameters)

    spreadsheet_informations += process_hierarchical_table(sheet, 84, 4, 10)
  end

  def process_simple_tables(sheet, simple_table_parameters)
    result = []

    simple_table_parameters.each do |params|
      result += process_simple_table(sheet, params[:row_start_position], params[:number_of_rows])
    end

    result
  end

  def serialize_value(cell)
    cell.to_s
  end

  def remove_none_key_value_pairs(hash)
    hash.reject { |k, v| k.empty? && v.empty? }
  end

  def process_simple_table(sheet, row_start_position, number_of_rows)
    records = []
    headers = sheet.row(row_start_position).map { |cell| serialize_value(cell) }

    first_content_row = row_start_position + 1
    last_content_row = first_content_row + number_of_rows

    (first_content_row..last_content_row).each do |i|
      row = sheet.row(i)
      values = row.map { |cell| serialize_value(cell) }
      record = Hash[headers.zip(values)]
      records << remove_none_key_value_pairs(record)
    end

    records
  end

  def calculate_num_leading_space_per_level(row_headers)
    row_headers.each_cons(2) do |current_header, next_header|
      current_spaces = current_header[/\A\s*/].length
      next_spaces = next_header[/\A\s*/].length

      return next_spaces - current_spaces if next_spaces != current_spaces
    end

    0
  end

  def process_hierarchical_table(sheet, row_start_position, row_size, number_of_rows)
    row_headers = sheet.column(1)[1..-1].map { |cell| serialize_value(cell) }

    column_start_position = 3
    col_headers = cells_in_row_range(sheet, row_start_position, column_start_position, row_size)

    first_content_row = row_start_position + 1
    last_content_row = first_content_row + number_of_rows

    num_leading_space_per_level = calculate_num_leading_space_per_level(row_headers)
    num_leading_space_per_level = 1 if num_leading_space_per_level.zero?

    processed_table = {}
    nodes = []

    (first_content_row..last_content_row).each do |i|
      row = sheet.row(i)
      cell_value = serialize_value(row[0])
      level = (cell_value.length - cell_value.lstrip.length) / num_leading_space_per_level
      label = cell_value.strip
      data_cells = row[1..-1]

      nodes = nodes[0...level]
      nodes << label

      if data_cells.any? { |c| !c.nil? }
        processed_table = add_data(processed_table, nodes, col_headers, data_cells)
      end
    end

    [remove_none_key_value_pairs(processed_table)]
  end

  def add_data(processed_table, nodes, col_headers, data_cells)
    current_level = processed_table

    nodes[0...-1].each do |node|
      current_level[node] ||= {}
      current_level = current_level[node]
    end

    current_level[nodes[-1]] = Hash[col_headers.zip(data_cells.map { |d| d ? serialize_value(d) : nil })]

    processed_table
  end

  def cells_in_row_range(sheet, row_num, column_num, row_size)
    range = []

    start_col_num = column_num
    end_col_num = column_num + row_size - 1

    (start_col_num..end_col_num).map do |col_num|
      cell_value = sheet.cell(row_num, col_num)

      range << cell_value
    end

    range[1..-1]
  end
end