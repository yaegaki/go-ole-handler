package olehandler

import (
	"errors"
	"fmt"
	"sync"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type OleHandler struct {
	App      *ole.IUnknown
	Handle   *ole.IDispatch
	parent   *OleHandler
	children []*OleHandler

	closed     bool
	closedChan chan struct{}
	m          *sync.RWMutex
	wg         *sync.WaitGroup
	once       *sync.Once
}

func CreateRootOleHandler(programID string) (*OleHandler, error) {
	app, err := oleutil.CreateObject(programID)
	if err != nil {
		return nil, err
	}

	handle, err := app.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		app.Release()
		return nil, err
	}

	handler := &OleHandler{
		App:      app,
		Handle:   handle,
		children: []*OleHandler{},

		closed:     false,
		closedChan: make(chan struct{}),
		m:          new(sync.RWMutex),
		wg:         new(sync.WaitGroup),
		once:       new(sync.Once),
	}

	return handler, nil
}

func CreateOleHandler(parent *OleHandler, handle *ole.IDispatch) *OleHandler {
	if parent != nil {
		parent.wg.Add(1)
	}
	return &OleHandler{
		Handle:   handle,
		parent:   parent,
		children: []*OleHandler{},

		closed:     false,
		closedChan: make(chan struct{}),
		m:          new(sync.RWMutex),
		wg:         new(sync.WaitGroup),
		once:       new(sync.Once),
	}
}

func (o *OleHandler) Close() {
	o.once.Do(func() {
		o.m.Lock()

		close(o.closedChan)
		o.closed = true
		children := o.children
		o.children = []*OleHandler{}

		for _, child := range children {
			child.Close()
		}

		o.m.Unlock()

		o.wg.Wait()
		o.Handle.Release()

		if o.parent != nil {
			if len(o.parent.children) != 0 {
				o.parent.m.Lock()
				// remove self from parent
				children := o.parent.children
				if children != nil {
					for i, child := range children {
						if child == o {
							last := len(children) - 1
							children[i] = children[last]
							o.parent.children = children[:last]
							break
						}
					}
				}
				o.parent.m.Unlock()
			}
			o.parent.wg.Done()
		}

		if o.App != nil {
			o.App.Release()
		}
	})
}

func (o *OleHandler) Closed() <-chan struct{} {
	return o.closedChan
}

func (o *OleHandler) SafeAccess(fn func() error) error {
	o.m.RLock()
	defer o.m.RUnlock()
	if o.closed {
		return errors.New("OleHandler is already closed.")
	}

	return fn()
}

func (o *OleHandler) GetOleHandler(property string) (handler *OleHandler, err error) {
	err = o.GetOleHandlerWithCallbackAndArgs(property, func(h *OleHandler) error {
		handler = h
		return nil
	})

	return handler, err
}

func (o *OleHandler) GetOleHandlerWithArgs(property string, args ...interface{}) (handler *OleHandler, err error) {
	err = o.GetOleHandlerWithCallbackAndArgs(property, func(h *OleHandler) error {
		handler = h
		return nil
	}, args...)

	return handler, err
}

func (o *OleHandler) GetOleHandlerWithCallback(property string, fn func(*OleHandler) error) error {
	return o.GetOleHandlerWithCallbackAndArgs(property, fn)
}

func (o *OleHandler) GetOleHandlerWithCallbackAndArgs(property string, fn func(*OleHandler) error, args ...interface{}) error {
	return o.SafeAccess(func() error {
		v, err := o.Handle.GetProperty(property, args...)
		if err != nil {
			return err
		}
		handle := v.ToIDispatch()
		if handle == nil {
			return errors.New(fmt.Sprintf("%v is not handle.", property))
		}

		handler := CreateOleHandler(o, handle)
		err = fn(handler)
		if err != nil {
			handler.Close()
			return err
		}

		o.children = append(o.children, handler)
		return nil
	})
}

func (o *OleHandler) GetProperty(property string, args ...interface{}) (result *ole.VARIANT, err error) {
	o.SafeAccess(func() error {
		result, err = o.Handle.GetProperty(property, args...)
		return err
	})

	return result, err
}

func (o *OleHandler) GetIntProperty(property string, args ...interface{}) (int, error) {
	v, err := o.GetProperty(property, args...)
	if err != nil {
		return 0, err
	}

	return int(v.Val), nil
}

func (o *OleHandler) GetStringProperty(property string, args ...interface{}) (string, error) {
	v, err := o.GetProperty(property, args...)
	if err != nil {
		return "", err
	}

	return v.ToString(), nil
}

func (o *OleHandler) GetBoolProperty(property string, args ...interface{}) (bool, error) {
	v, err := o.GetProperty(property, args...)
	if err != nil {
		return false, err
	}

	return v.Value().(bool), nil
}

func (o *OleHandler) PutProperty(property string, args ...interface{}) error {
	return o.SafeAccess(func() error {
		_, err := o.Handle.PutProperty(property, args...)
		return err
	})
}

func (o *OleHandler) CallMethod(property string, args ...interface{}) error {
	return o.SafeAccess(func() error {
		_, err := o.Handle.CallMethod(property, args...)
		return err
	})
}
