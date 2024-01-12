require 'fileutils'
require 'roo'
require 'open-uri'

download_folder_name = 'COMPROVANTES'

def create_directory_if_not_exists(directory)
  FileUtils.mkdir_p(directory) unless File.exist?(directory)
end

if ARGV.length != 1
  puts 'Usage: ruby app.rb <directory>'
  exit(1)
end

directory_path = ARGV[0]

unless File.exist?(directory_path) && File.directory?(directory_path)
  puts 'Invalid directory path. Please provide a valid directory.'
  exit(1)
end

Dir.entries(directory_path).each do |file|
  next unless File.extname(file) == '.xlsx'

  workbook = Roo::Spreadsheet.open File.join(directory_path, file)
  sheet = workbook.sheet 0

  required_headers = {
    comprovante: 'Comprovante Transfeera',
    favorecido: 'Favorecido',
    lote: 'Nome do lote',
    valor: 'Valor',
    status: 'Status',
    cpf: 'CPF ou CNPJ',
    email: 'Email'
  }

  current_headers = sheet.row(1)

  unless (current_headers & required_headers.values).size === required_headers.values.size
    puts "O arquivo XLSX não possui as colunas requeridas: #{file}"
    next
  end

  sheet.each(required_headers).with_index do |row_data, idx|    
    if idx === 0 then next end

    if row_data.all? { |k,v| !v.nil? }
      puts "Downloading PDF files from file '#{file}'..."
      devolvido_rows = []

      download_folder_path = File.join(directory_path, download_folder_name)
      create_directory_if_not_exists(download_folder_path)

      if row_data[:status] == 'DEVOLVIDO'
        devolvido_rows.push([
          row_data[:favorecido],
          row_data[:cpf],
          row_data[:email],
          '',
          '',
          '',
          '',
          '',
          row_data[:valor],
          '',
          '',
          row_data[:lote]
        ])
        next # Skip because there is no comprovante
      end

      # Generate a filename based on the extracted values
      filename = "PGM #{row_data[:favorecido]} #{row_data[:lote]} (#{row_data[:valor]}).pdf"

      # Download the PDF file
      begin
        pdf_path = File.join(download_folder_path, file.split('.')[0], filename)
        create_directory_if_not_exists(File.join(download_folder_path, file.split('.')[0]))

        # Save the PDF file
        open(pdf_path, 'wb') do |file|
          file << URI.parse(row_data[:comprovante]).open.read
        end

        puts "Downloaded and saved: #{pdf_path}"
      rescue StandardError => e
        puts "Error downloading PDF from #{row_data[:comprovante]}: #{e.message}"
      end
    else
      puts "One or more required columns not found in file '#{file}'."
    end

    if devolvido_rows.any?
      output_workbook = RubyXL::Workbook.new
      output_sheet = output_workbook[0]

      header_row = [
        [
          'Mantenha sempre o cabeçalho original da planilha e esta linha, mantendo os titulos e a ordem dos campos'
        ],
        [
          'Nome ou Razão Social',
          'CPF ou CNPJ',
          'Email (opcional)',
          'Banco',
          'Agência',
          'Conta',
          'Dígito da conta',
          'Tipo de Conta (Corrente ou Poupança)',
          'Valor',
          'ID integração (opcional)',
          'Data de agendamento (opcional)',
          'Descrição Pix (opcional)'
        ],
        *devolvido_rows
      ]

      header_row.each_with_index do |row, i|
        row.each_with_index do |value, j|
          output_sheet.add_cell(i, j, value)
        end
      end

      output_file_path = File.join(download_folder_path, "PGM DEVOLVIDOS - #{file}.xlsx")
      output_workbook.write(output_file_path)
      puts "New Excel file created pgm devolvidos: #{output_file_path}"
    else
      puts 'No rows with "DEVOLVIDO" status found.'
    end
  end
end