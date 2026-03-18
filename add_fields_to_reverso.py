import re

with open('/home/dev/credencializacion/CredencializacionFront/src/app/content/familiar/familiar.component.html', 'r') as f:
    text = f.read()

new_fields = """                                        <!-- CÓDIGO QR EN EL REVERSO -->
                                        <!-- En modo impresión, top: 11%, en no-print top: 12% para calzar el recuadro blanco -->
                                        <div class="campo-qr-reverso" style="position: absolute; top: 12%; left: 34%; width: 33%; aspect-ratio: 1 / 1; display: flex; align-items: center; justify-content: center; ">
                                                <img *ngIf="qrCodeDataUrl" 
                                                         [src]="qrCodeDataUrl" 
                                                         style="max-width: 100%; max-height: 100%; width: auto; height: auto; object-fit: contain;"
                                                         alt="QR Code">
                                                <div *ngIf="!qrCodeDataUrl && !isPrintMode"
                                                         class="text-muted text-center"
                                                         style="font-size: 0.6rem;">
                                                        <i class="fas fa-qrcode"></i>
                                                        <br>QR
                                                </div>
                                        </div>

                                        <!-- VIVIENDA -->
                                        <div class="campo-vivienda" style="position: absolute; top: 35%; left: 7%; width: 86%;">
                                                <ng-container *ngIf="isPrintMode">
                                                        <span style="color: #000000; font-size: 14px; font-family: 'NotoSans-Bold', Arial, sans-serif; display: block; text-align: left;">VIVIENDA: <span style="font-family: 'NotoSans', Arial, sans-serif;">{{ empleado.vivienda }}</span></span>
                                                </ng-container>
                                                <ng-container *ngIf="!isPrintMode">
                                                        <input type="text" [disabled]="!editable" class="campo-credencial" style="color: #000000; font-size: 0.7rem; border-bottom: 1px solid #ccc; width: 100%; padding-left: 5px;" [(ngModel)]="empleado.vivienda" placeholder="Vivienda">
                                                </ng-container>
                                        </div>

                                        <!-- MARCA -->
                                        <div class="campo-marca" style="position: absolute; top: 41%; left: 7%; width: 40%;">
                                                <ng-container *ngIf="isPrintMode">
                                                        <span style="color: #000000; font-size: 14px; font-family: 'NotoSans-Bold', Arial, sans-serif; display: block; text-align: left;">MARCA: <span style="font-family: 'NotoSans', Arial, sans-serif;">{{ empleado.marca }}</span></span>
                                                </ng-container>
                                                <ng-container *ngIf="!isPrintMode">
                                                        <input type="text" [disabled]="!editable" class="campo-credencial" style="color: #000000; font-size: 0.7rem; border-bottom: 1px solid #ccc; width: 100%; padding-left: 5px;" [(ngModel)]="empleado.marca" placeholder="Marca Vehículo">
                                                </ng-container>
                                        </div>

                                        <!-- MODELO -->
                                        <div class="campo-modelo" style="position: absolute; top: 41%; right: 7%; width: 40%;">
                                                <ng-container *ngIf="isPrintMode">
                                                        <span style="color: #000000; font-size: 14px; font-family: 'NotoSans-Bold', Arial, sans-serif; display: block; text-align: left;">MODELO: <span style="font-family: 'NotoSans', Arial, sans-serif;">{{ empleado.modelo }}</span></span>
                                                </ng-container>
                                                <ng-container *ngIf="!isPrintMode">
                                                        <input type="text" [disabled]="!editable" class="campo-credencial" style="color: #000000; font-size: 0.7rem; border-bottom: 1px solid #ccc; width: 100%; padding-left: 5px;" [(ngModel)]="empleado.modelo" placeholder="Modelo Vehículo">
                                                </ng-container>
                                        </div>

                                        <!-- COLOR -->
                                        <div class="campo-color" style="position: absolute; top: 47%; left: 7%; width: 40%;">
                                                <ng-container *ngIf="isPrintMode">
                                                        <span style="color: #000000; font-size: 14px; font-family: 'NotoSans-Bold', Arial, sans-serif; display: block; text-align: left;">COLOR: <span style="font-family: 'NotoSans', Arial, sans-serif;">{{ empleado.color }}</span></span>
                                                </ng-container>
                                                <ng-container *ngIf="!isPrintMode">
                                                        <input type="text" [disabled]="!editable" class="campo-credencial" style="color: #000000; font-size: 0.7rem; border-bottom: 1px solid #ccc; width: 100%; padding-left: 5px;" [(ngModel)]="empleado.color" placeholder="Color Vehículo">
                                                </ng-container>
                                        </div>

                                        <!-- PLACAS -->
                                        <div class="campo-placas" style="position: absolute; top: 47%; right: 7%; width: 40%;">
                                                <ng-container *ngIf="isPrintMode">
                                                        <span style="color: #000000; font-size: 14px; font-family: 'NotoSans-Bold', Arial, sans-serif; display: block; text-align: left;">PLACAS: <span style="font-family: 'NotoSans', Arial, sans-serif;">{{ empleado.placas }}</span></span>
                                                </ng-container>
                                                <ng-container *ngIf="!isPrintMode">
                                                        <input type="text" [disabled]="!editable" class="campo-credencial" style="color: #000000; font-size: 0.7rem; border-bottom: 1px solid #ccc; width: 100%; padding-left: 5px;" [(ngModel)]="empleado.placas" placeholder="Placas Vehículo">
                                                </ng-container>
                                        </div>

                                        <!-- CURP  -->"""

text = re.sub(r'(\s+<!-- CURP  -->)', '\n' + new_fields, text, count=1)

with open('/home/dev/credencializacion/CredencializacionFront/src/app/content/familiar/familiar.component.html', 'w') as f:
    f.write(text)

