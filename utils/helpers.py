def limpiar_serial(col):
    """Limpieza estándar de seriales — aplicar siempre antes de cualquier comparación."""
    return (
        col.astype(str)
           .str.upper()
           .str.strip()
           .str.replace(r'\s+', '', regex=True)
           .str.replace(r'\.0$', '', regex=True)
    )