function obtenerVistaUnidades() {
  return `
    <button class="btn-volver" onclick="ir('portal')" style="position: absolute; top: 20px; left: 20px; display: flex; align-items: center; gap: 4px;">
      <span class="material-symbols-outlined" style="font-size: 18px;">arrow_back</span> VOLVER
    </button>
    
    <div class="card-icon" style="margin-top: 30px;">
      <span class="material-symbols-outlined" style="font-size: 60px; color: #1c4587;">domain</span>
    </div>
    <div class="titulo-azul">Unidades Asistenciales</div>
    
    <div style="background-color: #fff8e1; border: 1px solid #ffe082; border-left: 5px solid #f29900; padding: 25px; text-align: left; margin-top: 30px; margin-bottom: 20px; border-radius: 6px;">
      <div style="display: flex; align-items: center; gap: 10px; color: #d68100; font-weight: bold; font-size: 18px; margin-bottom: 12px;">
        <span class="material-symbols-outlined" style="font-size: 28px;">construction</span> 
        MÓDULO EN CONSTRUCCIÓN
      </div>
      <p style="color: #555; font-size: 15px; margin: 0; line-height: 1.6;">
        Esta nueva aplicación se encuentra actualmente en etapa de desarrollo.<br><br>
        <strong>Nota de sistema:</strong> Este cartel ha sido inyectado con éxito desde el servidor externo (Unidades.gs).
      </p>
    </div>
  `;
}