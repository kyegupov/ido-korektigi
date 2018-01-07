regexalaReguli = [
    {de:/([0-9]{1,2}) di (januaro|februaro|marto|aprilo|mayo|junio|julio|agosto|septembro|oktobro|novembro|decembro)/, a:"$1 ma di $2"},
    {de:/([0-9]+)a yari/, a:"yari $1ma"},
    {de:/esas ([a-z]+)[ai]ta/, a:"$1esas"},
    {de:/esis ([a-z]+)[ai]ta/, a:"$1esis"},
    {de:/an sk/, a:"ana sk"},
    {de:/\baverajal/, a:"mezavalor"},
    {de:/\bciviten/, a:"civitan"},
    {de:/\bdirekt/, a:"diret"},
    {de:/\bkondemn/, a:"kondamn"},
    {de:/\bkonferenc/, a:"konfer"},
    {de:/\bmenac/, a:"minac"},
    {de:/\bmesur/, a:"mezur"},
    {de:/\bnedependes/, a:"nedepend"},
    {de:/\brelacion/, a:"relat"},
    {de:/\breprizent/, a:"reprezent"},
    {de:/\bterorigant/, a:"terorist"}   
];