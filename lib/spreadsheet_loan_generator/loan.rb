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

    def non_deferred_duration
      duration - total_deferred_duration
    end

    def total_deferred_duration
      deferred_and_capitalized + deferred
    end
  end
end
