package xlst

// options is the private structure that contains all the applied rendering options.
type options struct {
	unescapeHTML bool
}

// Option is the interface of functor that applies options to the Render method.
type Option func(o *options)

// OptionUnescapeHTML remove '&quot;' and other specific HTML strings from the result of substitution.
func OptionUnescapeHTML(o *options) {
	o.unescapeHTML = true
}
