module github.com/sunreaver/docanalysis

go 1.12

replace (
	code.sajari.com/docconv v1.1.0 => github.com/sunreaver/docconv v1.1.101
	github.com/extrame/xls => github.com/sunreaver/xls v1.0.1
	github.com/otiai10/gosseract => github.com/otiai10/gosseract v0.0.0-20190914143506-cca9300a9131
)

require (
	code.sajari.com/docconv v1.1.0
	github.com/JalfResi/justext v0.0.0-20170829062021-c0282dea7198 // indirect
	github.com/PuerkitoBio/goquery v1.5.1 // indirect
	github.com/advancedlogic/GoOse v0.0.0-20191112112754-e742535969c1 // indirect
	github.com/araddon/dateparse v0.0.0-20200409225146-d820a6159ab1 // indirect
	github.com/extrame/goyymmdd v0.0.0-20181026012948-914eb450555b // indirect
	github.com/extrame/ole2 v0.0.0-20160812065207-d69429661ad7 // indirect
	github.com/extrame/xls v0.0.0-00010101000000-000000000000
	github.com/golang/protobuf v1.4.2 // indirect
	github.com/jaytaylor/html2text v0.0.0-20200412013138-3577fbdbcff7 // indirect
	github.com/levigross/exp-html v0.0.0-20120902181939-8df60c69a8f5 // indirect
	github.com/olekukonko/tablewriter v0.0.4 // indirect
	github.com/pixiv/go-libjpeg v0.0.0-20190822045933-3da21a74767d
	github.com/unidoc/unidoc v2.2.0+incompatible
	github.com/unidoc/unioffice v1.4.0
	golang.org/x/image v0.0.0-20200801110659-972c09e46d76 // indirect
)
