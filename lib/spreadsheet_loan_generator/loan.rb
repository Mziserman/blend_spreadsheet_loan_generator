module SpreadsheetLoanGenerator
  class Loan
    attr_accessor :amount,
                  :duration,
                  :period_duration,
                  :rate,
                  :due_on,
                  :deferred_and_capitalized,
                  :deferred,
                  :type,
                  :interests_type

    def initialize(
      amount:,
      duration:,
      period_duration:,
      rate:,
      due_on:,
      deferred_and_capitalized:,
      deferred:,
      type:,
      interests_type:)
      @amount = amount.to_f
      @duration = duration.to_i
      @period_duration = period_duration.to_i
      @rate = rate.to_f
      @due_on = due_on.is_a?(Date) ? due_on : Date.parse(due_on)
      @deferred_and_capitalized = deferred_and_capitalized.to_i
      @deferred = deferred.to_i
      @type = type
      @interests_type = interests_type
    end

    def name_type
      return 'bullet' if bullet?
      return 'in_fine' if in_fine?

      type
    end

    def name_period_duration
      if period_duration.in?([1, 3, 6, 12])
        {
          '1' => 'month',
          '3' => 'quarter',
          '6' => 'semester',
          '12' => 'year'
        }[period_duration.to_s]
      else
        period_duration.to_s
      end
    end

    def name_deferred
      return '0' if fully_deferred?

      total_deferred_duration
    end

    def name_due_on
      due_on.strftime('%Y%m%d')
    end

    def name
      [
        name_type,
        name_period_duration,
        amount,
        (rate * 100).to_s,
        duration.to_s,
        name_deferred,
        name_due_on
      ].join('_')
    end

    def fully_deferred?
      duration > 1 && non_deferred_duration == 1
    end

    def bullet?
      fully_deferred? && deferred_and_capitalized == total_deferred_duration
    end

    def in_fine?
      fully_deferred? && deferred == total_deferred_duration
    end

    def non_deferred_duration
      duration - total_deferred_duration
    end

    def total_deferred_duration
      deferred_and_capitalized + deferred
    end
  end
end
