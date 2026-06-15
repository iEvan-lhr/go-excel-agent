package ops

import (
	"fmt"
	"sort"
	"strings"
)

type Registry struct {
	byOp map[string]OperationSpec
}

func NewRegistry(specs ...OperationSpec) *Registry {
	r := &Registry{byOp: make(map[string]OperationSpec)}
	for _, spec := range specs {
		r.Register(spec)
	}
	return r
}

func BuiltinRegistry() *Registry {
	return NewRegistry(BuiltinSpecs()...)
}

func (r *Registry) Register(spec OperationSpec) {
	if r.byOp == nil {
		r.byOp = make(map[string]OperationSpec)
	}
	op := strings.ToLower(strings.TrimSpace(spec.Op))
	if op == "" {
		return
	}
	spec.Op = op
	r.byOp[op] = spec
}

func (r *Registry) Get(op string) (OperationSpec, bool) {
	if r == nil || r.byOp == nil {
		return OperationSpec{}, false
	}
	spec, ok := r.byOp[strings.ToLower(strings.TrimSpace(op))]
	return spec, ok
}

func (r *Registry) MustGet(op string) (OperationSpec, error) {
	spec, ok := r.Get(op)
	if !ok {
		return OperationSpec{}, fmt.Errorf("operation spec not found: %s", op)
	}
	return spec, nil
}

func (r *Registry) List() []OperationSpec {
	if r == nil {
		return nil
	}
	out := make([]OperationSpec, 0, len(r.byOp))
	for _, spec := range r.byOp {
		out = append(out, spec)
	}
	sort.Slice(out, func(i, j int) bool {
		return out[i].Op < out[j].Op
	})
	return out
}
