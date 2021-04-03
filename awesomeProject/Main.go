package main

import (
	"encoding/json"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
)

type Clientes struct {
	Cliente string
	CodigoSistemaXYZ string
	Contas []string
}

func main() {

	f, err := excelize.OpenFile(".\\Exercicio.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	valor1, err := f.GetCellValue("DePara", "A2")
	valor2, _ := f.GetCellValue("DePara", "A3")
	valor3, _ := f.GetCellValue("DePara", "A4")
	valor4, _ := f.GetCellValue("DePara", "A5")

	var CodigoSistemaA string
	var CodigoSistemaB string
	var CodigoSistemaC string
	var CodigoSistemaD string

	switch {
	case valor1 == "A":
		CodigoSistemaA, _ = f.GetCellValue("DePara", "B2")
		break
	case valor1 == "B" :
		CodigoSistemaB, _ = f.GetCellValue("DePara", "B2")
		break
	case valor1 == "C" :
		CodigoSistemaC, _ = f.GetCellValue("DePara", "B2")
		break
	case valor1 == "D" :
		CodigoSistemaD, _ = f.GetCellValue("DePara", "B2")
		break
	}
	switch {
	case valor2 == "A":
		CodigoSistemaA, _ = f.GetCellValue("DePara", "B3")
		break
	case valor2 == "B" :
		CodigoSistemaB, _ = f.GetCellValue("DePara", "B3")
		break
	case valor2 == "C" :
		CodigoSistemaC, _ = f.GetCellValue("DePara", "B3")
		break
	case valor2 == "D" :
		CodigoSistemaD, _ = f.GetCellValue("DePara", "B3")
		break
	}
	switch {
	case valor3 == "A":
		CodigoSistemaA, _ = f.GetCellValue("DePara", "B4")
		break
	case valor3 == "B" :
		CodigoSistemaB, _ = f.GetCellValue("DePara", "B4")
		break
	case valor3 == "C" :
		CodigoSistemaC, _ = f.GetCellValue("DePara", "B4")
		break
	case valor3 == "D" :
		CodigoSistemaD, _ = f.GetCellValue("DePara", "B4")
		break
	}
	switch {
	case valor4 == "A":
		CodigoSistemaA, _ = f.GetCellValue("DePara", "B5")
		break
	case valor4 == "B" :
		CodigoSistemaB, _ = f.GetCellValue("DePara", "B5")
		break
	case valor4 == "C" :
		CodigoSistemaC, _ = f.GetCellValue("DePara", "B5")
		break
	case valor4 == "D" :
		CodigoSistemaD, _ = f.GetCellValue("DePara", "B5")
		break
	}

	tipoCliente1, err := f.GetCellValue("Contas", "A2")
	tipoCliente2, _ := f.GetCellValue("Contas", "A3")
	tipoCliente3, _ := f.GetCellValue("Contas", "A4")
	tipoCliente4, _ := f.GetCellValue("Contas", "A5")
	tipoCliente5, _ := f.GetCellValue("Contas", "A6")
	tipoCliente6, _ := f.GetCellValue("Contas", "A7")
	tipoCliente7, _ := f.GetCellValue("Contas", "A8")
	tipoCliente8, _ := f.GetCellValue("Contas", "A9")
	tipoCliente9, _ := f.GetCellValue("Contas", "A10")
	tipoCliente10, _ := f.GetCellValue("Contas", "A11")
	tipoCliente11, _ := f.GetCellValue("Contas", "A12")
	tipoCliente12, _ := f.GetCellValue("Contas", "A13")
	tipoCliente13, _ := f.GetCellValue("Contas", "A14")
	tipoCliente14, _ := f.GetCellValue("Contas", "A15")
	tipoCliente15, _ := f.GetCellValue("Contas", "A16")
	tipoCliente16, _ := f.GetCellValue("Contas", "A17")
	tipoCliente17, _ := f.GetCellValue("Contas", "A18")
	tipoCliente18, _ := f.GetCellValue("Contas", "A19")
	tipoCliente19, _ := f.GetCellValue("Contas", "A20")
	tipoCliente20, _ := f.GetCellValue("Contas", "A21")
	tipoCliente21, _ := f.GetCellValue("Contas", "A22")
	tipoCliente22, _ := f.GetCellValue("Contas", "A23")

	var ContasA []string
	var ContasB []string
	var ContasC []string
	var ContasD []string

	switch {
	case tipoCliente1 == "A":
		opt1, _ := f.GetCellValue("Contas", "B2")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente1 == "B":
		opt1, _ := f.GetCellValue("Contas", "B2")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente1 == "C":
		opt1, _ := f.GetCellValue("Contas", "B2")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente1 == "D":
		opt1, _ := f.GetCellValue("Contas", "B2")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente2 == "A":
		opt1, _ := f.GetCellValue("Contas", "B3")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente2 == "B":
		opt1, _ := f.GetCellValue("Contas", "B3")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente2 == "C":
		opt1, _ := f.GetCellValue("Contas", "B3")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente2 == "D":
		opt1, _ := f.GetCellValue("Contas", "B3")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente3 == "A":
		opt1, _ := f.GetCellValue("Contas", "B4")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente3 == "B":
		opt1, _ := f.GetCellValue("Contas", "B4")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente3 == "C":
		opt1, _ := f.GetCellValue("Contas", "B4")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente3 == "D":
		opt1, _ := f.GetCellValue("Contas", "B4")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente4 == "A":
		opt1, _ := f.GetCellValue("Contas", "B5")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente4== "B":
		opt1, _ := f.GetCellValue("Contas", "B5")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente4 == "C":
		opt1, _ := f.GetCellValue("Contas", "B5")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente4 == "D":
		opt1, _ := f.GetCellValue("Contas", "B5")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente5 == "A":
		opt1, _ := f.GetCellValue("Contas", "B6")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente5 == "B":
		opt1, _ := f.GetCellValue("Contas", "B6")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente5 == "C":
		opt1, _ := f.GetCellValue("Contas", "B6")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente5 == "D":
		opt1, _ := f.GetCellValue("Contas", "B6")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente6 == "A":
		opt1, _ := f.GetCellValue("Contas", "B7")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente6 == "B":
		opt1, _ := f.GetCellValue("Contas", "B7")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente6 == "C":
		opt1, _ := f.GetCellValue("Contas", "B7")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente6 == "D":
		opt1, _ := f.GetCellValue("Contas", "B7")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente7 == "A":
		opt1, _ := f.GetCellValue("Contas", "B8")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente7 == "B":
		opt1, _ := f.GetCellValue("Contas", "B8")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente7 == "C":
		opt1, _ := f.GetCellValue("Contas", "B8")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente7 == "D":
		opt1, _ := f.GetCellValue("Contas", "B8")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente8 == "A":
		opt1, _ := f.GetCellValue("Contas", "B9")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente8 == "B":
		opt1, _ := f.GetCellValue("Contas", "B9")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente8 == "C":
		opt1, _ := f.GetCellValue("Contas", "B9")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente8 == "D":
		opt1, _ := f.GetCellValue("Contas", "B9")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente9 == "A":
		opt1, _ := f.GetCellValue("Contas", "B10")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente9 == "B":
		opt1, _ := f.GetCellValue("Contas", "B10")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente9 == "C":
		opt1, _ := f.GetCellValue("Contas", "B10")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente9 == "D":
		opt1, _ := f.GetCellValue("Contas", "B10")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente10 == "A":
		opt1, _ := f.GetCellValue("Contas", "B11")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente10 == "B":
		opt1, _ := f.GetCellValue("Contas", "B11")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente10 == "C":
		opt1, _ := f.GetCellValue("Contas", "B11")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente10 == "D":
		opt1, _ := f.GetCellValue("Contas", "B11")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente11 == "A":
		opt1, _ := f.GetCellValue("Contas", "B12")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente11 == "B":
		opt1, _ := f.GetCellValue("Contas", "B12")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente11 == "C":
		opt1, _ := f.GetCellValue("Contas", "B12")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente11 == "D":
		opt1, _ := f.GetCellValue("Contas", "B12")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente12 == "A":
		opt1, _ := f.GetCellValue("Contas", "B13")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente12 == "B":
		opt1, _ := f.GetCellValue("Contas", "B13")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente12 == "C":
		opt1, _ := f.GetCellValue("Contas", "B13")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente12 == "D":
		opt1, _ := f.GetCellValue("Contas", "B13")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente13 == "A":
		opt1, _ := f.GetCellValue("Contas", "B14")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente13 == "B":
		opt1, _ := f.GetCellValue("Contas", "B14")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente13 == "C":
		opt1, _ := f.GetCellValue("Contas", "B14")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente13 == "D":
		opt1, _ := f.GetCellValue("Contas", "B14")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente14 == "A":
		opt1, _ := f.GetCellValue("Contas", "B15")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente14 == "B":
		opt1, _ := f.GetCellValue("Contas", "B15")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente14 == "C":
		opt1, _ := f.GetCellValue("Contas", "B15")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente14 == "D":
		opt1, _ := f.GetCellValue("Contas", "B15")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente15 == "A":
		opt1, _ := f.GetCellValue("Contas", "B16")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente15 == "B":
		opt1, _ := f.GetCellValue("Contas", "B16")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente15 == "C":
		opt1, _ := f.GetCellValue("Contas", "B16")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente15 == "D":
		opt1, _ := f.GetCellValue("Contas", "B16")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente16 == "A":
		opt1, _ := f.GetCellValue("Contas", "B17")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente16 == "B":
		opt1, _ := f.GetCellValue("Contas", "B17")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente16 == "C":
		opt1, _ := f.GetCellValue("Contas", "B17")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente16 == "D":
		opt1, _ := f.GetCellValue("Contas", "B17")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente17 == "A":
		opt1, _ := f.GetCellValue("Contas", "B18")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente17 == "B":
		opt1, _ := f.GetCellValue("Contas", "B18")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente17 == "C":
		opt1, _ := f.GetCellValue("Contas", "B18")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente17 == "D":
		opt1, _ := f.GetCellValue("Contas", "B18")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente18 == "A":
		opt1, _ := f.GetCellValue("Contas", "B19")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente18 == "B":
		opt1, _ := f.GetCellValue("Contas", "B19")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente18 == "C":
		opt1, _ := f.GetCellValue("Contas", "B19")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente18 == "D":
		opt1, _ := f.GetCellValue("Contas", "B19")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente19 == "A":
		opt1, _ := f.GetCellValue("Contas", "B20")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente19 == "B":
		opt1, _ := f.GetCellValue("Contas", "B20")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente19 == "C":
		opt1, _ := f.GetCellValue("Contas", "B20")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente19 == "D":
		opt1, _ := f.GetCellValue("Contas", "B20")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente20 == "A":
		opt1, _ := f.GetCellValue("Contas", "B21")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente20 == "B":
		opt1, _ := f.GetCellValue("Contas", "B21")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente20 == "C":
		opt1, _ := f.GetCellValue("Contas", "B21")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente20 == "D":
		opt1, _ := f.GetCellValue("Contas", "B21")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente21 == "A":
		opt1, _ := f.GetCellValue("Contas", "B22")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente21 == "B":
		opt1, _ := f.GetCellValue("Contas", "B22")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente21 == "C":
		opt1, _ := f.GetCellValue("Contas", "B22")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente21 == "D":
		opt1, _ := f.GetCellValue("Contas", "B22")
		ContasD = append(ContasD, opt1)
		break
	}
	switch {
	case tipoCliente22 == "A":
		opt1, _ := f.GetCellValue("Contas", "B23")
		ContasA = append(ContasA, opt1)
		break
	case tipoCliente22 == "B":
		opt1, _ := f.GetCellValue("Contas", "B23")
		ContasB = append(ContasB, opt1)
		break
	case tipoCliente22 == "C":
		opt1, _ := f.GetCellValue("Contas", "B23")
		ContasC = append(ContasC, opt1)
		break
	case tipoCliente22 == "D":
		opt1, _ := f.GetCellValue("Contas", "B23")
		ContasD = append(ContasD, opt1)
		break
	}


	clienteA := &Clientes{Cliente: "A",
		CodigoSistemaXYZ: CodigoSistemaA,
		Contas:           ContasA,
	}
	a, _ := json.MarshalIndent(clienteA, "", "")
	fmt.Println(string(a))

	clienteB := &Clientes{Cliente: "B",
		CodigoSistemaXYZ: CodigoSistemaB,
		Contas:           ContasB,
		}
		b, _ := json.MarshalIndent(clienteB, "", "")
		fmt.Println(string(b))

		clienteC := &Clientes{Cliente: "C",
			CodigoSistemaXYZ: CodigoSistemaC,
			Contas:           ContasC,
			}
			c, _ := json.MarshalIndent(clienteC, "", "")
			fmt.Println(string(c))

			clienteD := &Clientes{Cliente: "D",
				CodigoSistemaXYZ: CodigoSistemaD,
				Contas:           ContasD,
				}
				d, _ := json.MarshalIndent(clienteD, "", "")
				fmt.Println(string(d))

											}