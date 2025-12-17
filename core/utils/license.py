import logging
from pathlib import Path

import aspose.words as aw

logger = logging.getLogger(__name__)

def apply_license(license_path: str | None=None) -> None:
    if license_path and Path(license_path).exists():
        logger.info(f'Applying Aspose.Words license from: {license_path}')
        lic = aw.License()
        lic.set_license(license_path)
    else:
        logger.warning('No valid Aspose.Words license found. Running in Evaluation mode. Set ASPOSE_WORDS_LICENSE_PATH.')
