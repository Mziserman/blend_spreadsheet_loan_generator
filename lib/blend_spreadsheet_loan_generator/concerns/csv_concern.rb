module BlendSpreadsheetLoanGenerator
  module CsvConcern
    extend ActiveSupport::Concern

    included do
      def generate_csv(worksheet:, target_path:)
        filename = File.join(target_path, "#{loan.name}.csv")
        CSV.open(filename, 'wb') do |csv|
          loan.duration.times do |line|
            row = []
            columns.each.with_index do |name, column|
              row << (
                case name
                when 'index'
                  worksheet[line + 2, column + 1]
                when 'due_on'
                  Date.parse(worksheet[line + 2, column + 1]).strftime('%m/%d/%Y')
                else
                  worksheet[line + 2, column + 1].gsub(',', '.').to_f
                end
              )
            end
            csv << row
          end
        end
      end
    end
  end
end
