import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class CpfCnpj implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private _value: string = "";	
	private _notifyOutputChanged: () => void;
	private input: HTMLInputElement;
	private _context: ComponentFramework.Context<IInputs>;

	/**
	 * Empty constructor.
	 */
	constructor()
	{
	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		this._context = context;

		this.input = document.createElement("input");
		this.input.setAttribute("type", "text");
		this.input.addEventListener("blur", this.onInputBlur.bind(this));
		this.input.addEventListener("keydown", this.onInputKeyDown.bind(this));
		this._notifyOutputChanged = notifyOutputChanged;

		container.appendChild(this.input);
	}

	private onInputBlur(event: Event): void {
		let isValid = this.checkCPFCNPJ(this.input.value);	
		if (!isValid){
			this.input.value = "";
			this._value = "";
		}	
		this._notifyOutputChanged();
	}

	private onInputKeyDown(event: Event): void {	
		setTimeout(() => {
			this.fMascEx();
		}, 1);	   
	}

	private fMascEx() : void {	
		var fieldType = this._context.parameters.fieldType.raw!.toLowerCase();
		var isFieldTypeCPF = (!fieldType || fieldType == "" || fieldType == "cpf");
		this.input.setAttribute("maxlength", isFieldTypeCPF ? "14" : "18");
	
		if (isFieldTypeCPF)
			this._value = this.mCPF(this.input.value);
		else
			this._value = this.mCNPJ(this.input.value);
		
		this._notifyOutputChanged();
	}

	private mCPF(cpf:string) : string {
		cpf=cpf.replace(/\D/g,"");
		cpf=cpf.replace(/(\d{3})(\d)/,"$1.$2");
		cpf=cpf.replace(/(\d{3})(\d)/,"$1.$2");
		cpf=cpf.replace(/(\d{3})(\d{1,2})$/,"$1-$2");
		return cpf;
	}
	
	private mCNPJ(cnpj:string) : string {
		cnpj=cnpj.replace(/\D/g,"");
		cnpj=cnpj.replace(/^(\d{2})(\d)/,"$1.$2");
		cnpj=cnpj.replace(/^(\d{2})\.(\d{3})(\d)/,"$1.$2.$3");
		cnpj=cnpj.replace(/\.(\d{3})(\d)/,".$1/$2");
		cnpj=cnpj.replace(/(\d{4})(\d)/,"$1-$2");
		return cnpj;
	}
			
	private checkCPFCNPJ(txt:string) : boolean {
		var isValid = true;
		var msgError = "";
		var txtContent = txt.replace(/\./g, "").replace("-", "").replace("/", "").trim();

		if (txtContent == "") return true;

		if (txtContent.length == 14) {
			isValid = this.IsCNPJ(txtContent);
			if (!isValid) msgError = "CNPJ Inválido";
		} else {
			isValid = this.IsCPF(txtContent);
			if (!isValid) msgError = "CPF Inválido";
		}

		if (msgError != "") {
			this._context.navigation.openAlertDialog({
				confirmButtonLabel: "OK", 
				text: msgError
				});
		}

		return isValid;
	}

	private IsCNPJ(cnpj:string): boolean {
		cnpj = cnpj.replace(/[^\d]+/g, '');
		if (cnpj == '') return false;
		if (cnpj.length != 14) return false;

		// Elimina CNPJs invalidos conhecidos
		if (cnpj == "00000000000000" ||
			cnpj == "11111111111111" ||
			cnpj == "22222222222222" ||
			cnpj == "33333333333333" ||
			cnpj == "44444444444444" ||
			cnpj == "55555555555555" ||
			cnpj == "66666666666666" ||
			cnpj == "77777777777777" ||
			cnpj == "88888888888888" ||
			cnpj == "99999999999999")
			return false;

		// Valida DVs
		var tamanho:number;
		var numeros:string;
		var digitos:string;
		var soma:number;

		tamanho = cnpj.length - 2
		numeros = cnpj.substring(0, tamanho);
		digitos = cnpj.substring(tamanho);
		soma = 0;
		var pos:number = tamanho - 7;
		
		for (var i:number = tamanho; i >= 1; i--) {
			soma += ((numeros.charAt(tamanho - i) as any) as number) * pos--;
			if (pos < 2)
				pos = 9;
		}
		var resultado:number = soma % 11 < 2 ? 0 : 11 - soma % 11;
		if (resultado != ((digitos.charAt(0) as any) as number) )
			return false;

		tamanho = tamanho + 1;
		numeros = cnpj.substring(0, tamanho);
		soma = 0;
		pos = tamanho - 7;
		
		for (i = tamanho; i >= 1; i--) {
			soma += ((numeros.charAt(tamanho - i) as any) as number) * pos--;
			if (pos < 2)
				pos = 9;
		}
		resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
		if (resultado != ((digitos.charAt(1) as any) as number) )
			return false;

		return true;
	}

	private IsCPF(strCPF:string) : boolean {
		var Soma;
		var Resto;
		Soma = 0;
		// Elimina CPFs invalidos conhecidos
		if (strCPF == "00000000000" ||
			strCPF == "11111111111" ||
			strCPF == "22222222222" ||
			strCPF == "33333333333" ||
			strCPF == "44444444444" ||
			strCPF == "55555555555" ||
			strCPF == "66666666666" ||
			strCPF == "77777777777" ||
			strCPF == "88888888888" ||
			strCPF == "99999999999")
			return false;

		for (var i = 1; i <= 9; i++) Soma = Soma + parseInt(strCPF.substring(i - 1, i)) * (11 - i);
		Resto = (Soma * 10) % 11;

		if ((Resto == 10) || (Resto == 11)) Resto = 0;
		if (Resto != parseInt(strCPF.substring(9, 10))) return false;

		Soma = 0;
		for (var i = 1; i <= 10; i++) Soma = Soma + parseInt(strCPF.substring(i - 1, i)) * (12 - i);
		Resto = (Soma * 10) % 11;

		if ((Resto == 10) || (Resto == 11)) Resto = 0;
		if (Resto != parseInt(strCPF.substring(10, 11))) return false;
		return true;
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		// Add code to update control view
		this._value = context.parameters.cpfcnpjValue.raw!;
		this.input.value = this._value != null ? this._value: "";	
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {
			cpfcnpjValue : this._value
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}
}