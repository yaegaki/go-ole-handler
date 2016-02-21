# go-ole-handler
A utility for go-ole

Thread safe and Release safe.(If close top olehandler, close all children.)

## Sample
```go
package main

import (
	"log"

	"github.com/go-ole/go-ole"
	"github.com/yaegaki/go-ole-handler"
)

func main() {
	err := Itunes()
	if err != nil {
		log.Fatal(err)
	}
}

func Itunes() error {
	err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	if err != nil {
		return err
	}
	defer ole.CoUninitialize()

	handler, err := olehandler.CreateRootOleHandler("iTunes.Application")
	if err != nil {
		return err
	}
	defer handler.Close()

	track, err := handler.GetOleHandler("CurrentTrack")
	if err != nil {
		return err
	}

	name, err := track.GetStringProperty("Name")
	if err != nil {
		return err
	}

	artist, err := track.GetStringProperty("Artist")
	if err != nil {
		return err
	}

	log.Printf("%v %v", name, artist)

	return nil
}
```
