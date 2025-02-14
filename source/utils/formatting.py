# utils/formatting.py

import logging

logger = logging.getLogger(__name__)


def format_price(value, config):
    """
    Formats the price based on the specified type.

    Args:
        value (float or int): The price value to format.
        config (dict): Configuration object containing the 'type' key.

    Returns:
        str: The formatted price.

    Raises:
        ValueError: If the 'type' in config is invalid or value is not a number.
    """
    price_type = config.get('type')
    if price_type is None:
        logger.error("Configuration object missing 'type' key.")
        raise ValueError("Configuration object must contain a 'type' key.")

    try:
        numeric_value = float(value)
    except (ValueError, TypeError) as e:
        logger.error(f"Invalid value '{value}' for formatting: {e}")
        raise ValueError(f"Value '{value}' must be a number.") from e

    try:
        if price_type.lower() == 'ytm':
            formatted_value = f"{numeric_value:.2f}%"
        elif price_type.lower() == 'spread':
            formatted_value = f"{numeric_value:.1f}"
        elif price_type.lower() == 'price':
            formatted_value = f"${numeric_value:.3f}"
        else:
            logger.error(f"Invalid type '{price_type}' provided for formatting.")
            raise ValueError(f"Invalid type '{price_type}'. Expected 'price', 'ytm', or 'spread'.")
    except Exception as e:
        logger.error(f"Error formatting value '{value}' with type '{price_type}': {e}")
        raise ValueError(f"Error formatting value '{value}' with type '{price_type}'.") from e

    logger.debug(f"Formatted value: {formatted_value} (Type: {price_type})")
    return formatted_value

ALLOWED_SIZES = {1000, 10000, 100000, 250000, 500000,
                 1000000, 2000000, 3000000, 4000000, 5000000}

# Global cache to store previously computed results.
_quantity_cache = {}

def get_valid_quantity(value):
    """
    Convert the input value to a valid quantity.

    The value may be a numeric type or a string representing a number.
    If the converted value is one of the allowed sizes, it is returned.
    Otherwise, the closest valid quantity from ALLOWED_SIZES is returned.
    If the conversion fails, returns None.

    Uses caching so that repeated calls with the same value do not require
    re-computation.
    """
    # Check if the result is already cached.
    if value in _quantity_cache:
        return _quantity_cache[value]

    try:
        numeric_value = float(value)
        # Convert to an integer if the number is whole, otherwise round.
        if numeric_value.is_integer():
            numeric_value = int(numeric_value)
        else:
            numeric_value = round(numeric_value)
    except (ValueError, TypeError):
        _quantity_cache[value] = None
        return None

    if numeric_value in ALLOWED_SIZES:
        _quantity_cache[value] = numeric_value
        return numeric_value
    else:
        # Determine the closest allowed size.
        closest = min(ALLOWED_SIZES, key=lambda x: (abs(x - numeric_value), x))
        _quantity_cache[value] = closest
        return closest
