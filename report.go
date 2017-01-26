// Read from pam.csv, write HTML to stdout
package main

import (
	"encoding/csv"
	"html/template"
	"log"
	"os"
	"strings"
)

var sections = map[string]string{
	"Birth":       "Essentials",
	"Mother":      "Mother",
	"Father":      "Father",
	"doctor":      "Doctor",
	"development": "Development",
	"adults":      "Adults",
	"country":     "Country",
	"Allergy":     "Allergies",
}

var skipColumns = []string{"Timestamp"}

func findSection(s string, seen map[string]bool) (string, bool) {
	for k, v := range sections {
		if seen[k] {
			continue
		}
		if strings.Contains(s, k) {
			seen[k] = true
			return v, true
		}
	}
	return "", false
}

func main() {
	f, err := os.Open("pam.csv")
	if err != nil {
		log.Fatal(err)
	}
	r := csv.NewReader(f)
	records, err := r.ReadAll()
	if err != nil {
		log.Fatal(err)
	}
	skipColumnMap := make(map[string]bool)
	for _, k := range skipColumns {
		skipColumnMap[k] = true
	}

	students := splitStudents(records[0], records[1:], skipColumnMap)

	tmpl, err := template.New("test").Parse(outputTemplate)
	if err != nil {
		panic(err)
	}
	if err = tmpl.Execute(os.Stdout, students); err != nil {
		log.Fatal(err)
	}

}

const outputTemplate = `
<html>
<head><title>students</title>
<style media="print" type="text/css">
.pagebreak {page-break-before: always;}
</style>
<style media="all" type="text/css">
th {
	text-align:right;
	width: 25em;
	font-weight: normal;
}
</style>
</head>
<body>
{{range .}}
<h1 class="pagebreak">{{.LastName}}, {{.FirstName}}</h1>
{{range .Sections}}
<h2>{{.Name}}</h2>
<table>
{{- range .Data}}
<tr><th>{{.K}}:</th><td>{{.V}}</td></tr>
{{- end}}
</table>
{{end}}
{{end}}
</body>
</html>
`

type student struct {
	FirstName, LastName string
	Sections            []*section
}

type element struct {
	K, V string
}
type section struct {
	Name string
	Data []element
}

func splitStudents(headers []string, data [][]string, ignoreHeader map[string]bool) []student {
	var reply []student

	for _, d := range data {
		var stud student
		seen := make(map[string]bool)
		var s *section
		for i := range d {
			if ignoreHeader[headers[i]] {
				continue
			}
			if name, ok := findSection(headers[i], seen); ok {
				s = &section{Name: name}
				stud.Sections = append(stud.Sections, s)
			}
			if d[i] == "" {
				continue
			}
			switch headers[i] {
			case "Child's last name":
				stud.LastName = d[i]
			case "Child's first name":
				stud.FirstName = d[i]
			default:
				s.Data = append(s.Data, element{headers[i], d[i]})
			}
		}
		reply = append(reply, stud)
	}
	return reply
}
