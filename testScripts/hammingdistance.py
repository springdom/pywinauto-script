def hamming_distance(s1, s2):
    """Return the Hamming distance between equal-length sequences"""
    if len(s1) != len(s2):
        raise ValueError("Undefined for sequences of unequal length")
    
    hm_distance = sum(el1 != el2 for el1, el2 in zip(s1, s2))
    hm_distance -= 1
    if hm_distance <= 0:
        hm_distance = 0
    print(hm_distance)
    #return sum(el1 != el2 for el1, el2 in zip(s1, s2))

