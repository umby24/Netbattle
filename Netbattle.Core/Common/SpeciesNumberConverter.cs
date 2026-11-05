using Netbattle.Common;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Netbattle.Core.Common;

/// <summary>
/// Stores a species in team.jnb as its dex number, and links back to the database entry on load.
/// Older team.jnb files embed the whole species, so those are read by their "No" field.
/// </summary>
public class SpeciesNumberConverter : JsonConverter<PokemonSpecies> {
    public override void WriteJson(JsonWriter writer, PokemonSpecies? value, JsonSerializer serializer) {
        writer.WriteValue(value?.No ?? 0);
    }

    public override PokemonSpecies? ReadJson(JsonReader reader, Type objectType, PokemonSpecies? existingValue,
        bool hasExistingValue, JsonSerializer serializer) {
        JToken token = JToken.Load(reader);
        JToken? number = token.Type == JTokenType.Object ? token["No"] : token;

        if (number == null || number.Type != JTokenType.Integer)
            return null;

        return PokemonDatabase.BasePokemonMap.GetValueOrDefault(number.Value<int>());
    }
}
