package memory

type Options struct {
	ColumnTagger        ColumnTagger
	IntentGeneralizer   IntentGeneralizer
	ExecutionSummarizer ExecutionSummarizer
}

type Option func(*Options)

func WithColumnTagger(tagger ColumnTagger) Option {
	return func(options *Options) {
		options.ColumnTagger = tagger
	}
}

func WithIntentGeneralizer(generalizer IntentGeneralizer) Option {
	return func(options *Options) {
		options.IntentGeneralizer = generalizer
	}
}

func WithExecutionSummarizer(summarizer ExecutionSummarizer) Option {
	return func(options *Options) {
		options.ExecutionSummarizer = summarizer
	}
}

func applyOptions(options []Option) Options {
	var out Options
	for _, option := range options {
		if option != nil {
			option(&out)
		}
	}
	return out
}
